import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {runInNewContext} from 'node:vm';
const source=readFileSync(new URL('../server.mjs',import.meta.url),'utf8');
const fn=name=>source.match(new RegExp(`(?:async )?function ${name}\\([^]*?\\n}\\n`))[0];
const cleanSheetWhitespace=v=>String(v??'').trim();
const record={recordId:'source-handoff',sourceRecordId:'source',contentType:'EXC',sourceVideoFileId:'reviewed-video',
  sourceContentId:'VV-test',bookTitle:'Outside Press',bookShortener:'',author:'Poet',poemTitle:'Poem',excerpt:'Approved words',approvedAt:'approval-time'};
const fallback=runInNewContext(`(${fn('videoExcerptBookFallback')})`,{cleanSheetWhitespace});
const rejected={error:'no_importable_weaver_excerpt_records'};
test('only an explicit pre-content book mapping rejection can use the verified video source',()=>{
  const result=fallback(record,{status:400},rejected);
  assert.equal(result.payload.bookTitle,'');assert.equal(result.payload.sourceBookTitle,record.bookTitle);
  for(const key of ['recordId','sourceRecordId','sourceVideoFileId','sourceContentId','author','poemTitle','excerpt','approvedAt'])assert.equal(result.payload[key],record[key]);
  assert.equal(record.bookTitle,'Outside Press');assert.equal(result.evidence.receiverError,rejected.error);
});
test('timeouts, server errors, ordinary validation errors, partial results and success are never replayed',()=>{
  for(const [status,body]of [[200,rejected],[503,rejected],[400,{error:'timeout'}],[400,{...rejected,createdCount:1}],
    [400,{...rejected,updatedCount:1}],[400,{...rejected,results:[{id:'partial'}]}],[400,{}]])assert.equal(fallback(record,{status},body),null);
});
test('book imports and missing video identities cannot enter the fallback',()=>{
  for(const extra of [{contentType:'FP'},{sourceVideoFileId:''},{sourceContentId:''},{sourceRecordId:''},{bookTitle:''},{bookShortener:'KNOWN'}])assert.equal(fallback({...record,...extra},{status:400},rejected),null);
});
function sender(responses){const calls=[];const context={cleanSheetWhitespace,poetryPleaseApiKey:'test',poetryPleaseApiUrl:'https://receiver.example/api',
  buildPoetryPleaseExcerptRecord:r=>({...r}),syncWeaverRuntimeDb:async()=>({records:[]}),poetryPleaseVideoIdForSourceFile:()=> 'VV-test',
  fetch:async(_url,options)=>{calls.push(JSON.parse(options.body));const next=responses.shift();if(next instanceof Error)throw next;return {status:next.status,ok:next.status>=200&&next.status<300,json:async()=>next.body};}};
  const send=runInNewContext(fn('videoExcerptBookFallback')+'\n'+fn('handoffApprovedExcerptsToPoetryPlease')+';handoffApprovedExcerptsToPoetryPlease',context);return {send,calls};}
test('one safe fallback preserves the approved source and requires an actual item acknowledgement',async()=>{
  const {send,calls}=sender([{status:400,body:rejected},{status:200,body:{ok:true,createdCount:1,results:[{id:'EXC-test'}]}}]);
  const result=await send([record]);assert.equal(result.ok,true);assert.equal(calls.length,2);assert.equal(calls[0].payload.bookTitle,'Outside Press');assert.equal(calls[1].payload.bookTitle,'');assert.equal(result.results[0].poetryPleaseItemId,'EXC-test');assert.equal(result.results[0].routingEvidence.sourceBookTitle,'Outside Press');assert.equal(record.bookTitle,'Outside Press');
});
test('a failed fallback is held after two total requests and never recursively retried',async()=>{
  for(const second of [{status:400,body:rejected},new Error('timeout')]){const {send,calls}=sender([{status:400,body:rejected},second]);const result=await send([record]);assert.equal(result.ok,false);assert.equal(calls.length,2);}
});
test('a normal successful book import stays on its original path',async()=>{
  const {send,calls}=sender([{status:200,body:{ok:true,results:[{id:'BOOK-EXC'}]}}]);const result=await send([record]);assert.equal(calls.length,1);assert.equal(result.results[0].poetryPleaseItemId,'BOOK-EXC');assert.equal(result.results[0].routingEvidence,undefined);
});
