import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {runInNewContext} from 'node:vm';
import {excerptFingerprint,findExcerptAcknowledgement,verifyExcerptReadback} from '../excerpt_receipt.mjs';
const record={recordId:'handoff',sourceRecordId:'source',contentType:'EXC',author:'Poet',poemTitle:'Title',excerpt:'A line\nwith meaning.'};
const health={ok:true,ledgers:[{id:'ledger',status:'completed',outcomes:[{sourceRecordId:'source',contentId:'EXC-ONE',canonicalExcerptId:'EXC-ONE',canonicalVideoId:'VIDEO',outcome:'created'}]}]};
const body={imageId:'EXC-ONE',item:['Poet','Title','','EXC-ONE','','','','EXC','A line with meaning.']};
test('only exact source acknowledgements and current matching content verify a receipt',()=>{
 const ack=findExcerptAcknowledgement(record,health,'VIDEO');
 const receipt=verifyExcerptReadback(record,ack,body);
 assert.equal(receipt.verified,true);assert.equal(receipt.excerptFingerprint,excerptFingerprint(record.excerpt));
 assert.deepEqual(receipt.importLedgerIds,['ledger']);
 for(const changed of [{...body,imageId:'other'},{...body,item:[...body.item.slice(0,8),'Different line']},
   {...body,item:['Other poet',...body.item.slice(1)]},{...body,item:['Poet','Other title',...body.item.slice(2)]}]) {
   assert.throws(()=>verifyExcerptReadback(record,ack,changed));
 }
});
test('missing, conflicting, failed, wrong-video and changed saved identities remain held',()=>{
 assert.throws(()=>findExcerptAcknowledgement(record,{ok:true,ledgers:[]}),/does not prove absence/);
 assert.throws(()=>findExcerptAcknowledgement(record,health,'OTHER'));
 assert.throws(()=>findExcerptAcknowledgement({...record,poetryPleaseItemId:'OTHER'},health));
 const outcome=health.ledgers[0].outcomes[0];
 for(const h of [{ok:false,ledgers:[]},{ok:true,ledgers:[{id:'l',status:'failed',outcomes:[outcome]}]},
  {ok:true,ledgers:[{id:'l',status:'completed',outcomes:[outcome,{...outcome,contentId:'OTHER',canonicalExcerptId:'OTHER'}]}]}]) {
  assert.throws(()=>findExcerptAcknowledgement(record,h));
 }
});
const app=readFileSync(new URL('../public/app.js',import.meta.url),'utf8');
const next=runInNewContext(`(${app.match(/function videoPipelineNextStep\([^]*?\n}\n/)[0]})`);
test('next steps separate identity, editorial selection, copy state and excerpt approval',()=>{
 assert.match(next({gate:{}}),/editorial decision/);
 assert.match(next({publicationRestricted:true,gate:{}}),/clearance/);
 assert.match(next({sourceMediaType:'FV',requiresManualIdentityReview:true,gate:{decision:'ready_for_poetry_please'}}),/recording identity/);
 assert.match(next({sourceMediaType:'OM',gate:{decision:'ready_for_poetry_please'}}),/finished edit/);
 assert.match(next({gate:{poetryPleaseHandoff:{status:'sent_to_poetry_please',durableCopyStatus:'copying'}}}),/copying/);
 assert.match(next({gate:{poetryPleaseHandoff:{status:'sent_to_poetry_please',durableCopyStatus:'verified'}}}),/Excerpt approval.*separate/);
});
const server=readFileSync(new URL('../server.mjs',import.meta.url),'utf8');
test('reconciliation rereads receipt, persists and reads back without an import',async()=>{
 const calls=[];const context={cleanSheetWhitespace:x=>String(x||'').trim(),excerptFingerprint,findExcerptAcknowledgement,verifyExcerptReadback,
  poetryPleaseApiKey:'test',poetryPleaseApiUrl:'https://receiver.invalid/api',AbortSignal,
  getExcerptHandoffRecords:async()=>({records:[record]}),buildPoetryPleaseExcerptRecord:r=>r,invalidateQueueSnapshots(){},
  syncWeaverRuntimeDb:async(action,payload)=>{calls.push({action,payload});return action==='reconcile_excerpt_handoff'?{saved:true}
    :{records:[calls.some(c=>c.action==='reconcile_excerpt_handoff')?{...record,handoffStatus:'sent',poetryPleaseItemId:'EXC-ONE'}:record]};},
  fetch:async(url,options)=>{calls.push({url,options});return {ok:true,json:async()=>url.includes('weaverImportHealth')?health:body};}};
 const inspect=runInNewContext(`(${server.match(/async function inspectExcerptHandoffReceipt\([^]*?\n}\n/)[0]})`,context);
 const r=await inspect('handoff','operator@example.com',true);assert.equal(r.saved,true);
 assert.ok(calls.find(c=>c.action==='reconcile_excerpt_handoff').payload.receipt.verified);
 assert.ok(calls.filter(c=>c.url).every(c=>!c.options.method||c.options.method==='GET'));
 assert.equal(calls.filter(c=>c.action==='get_excerpt_handoffs').length,2);
 assert.match(server,/"\/api\/excerpts\/handoffs\/reconcile"/);
 assert.match(server.slice(server.indexOf('if (url.pathname === "/api/excerpts/handoffs/receipt"'),server.indexOf('if (url.pathname === "/api/excerpts/handoffs/retry"')),/verifyAdministrativeCaller/);
});
test('absent runtime ledger requires unchanged approved source before receipt-only creation',async()=>{
 for(const changed of [false,true]) {
  const calls=[];let reads=0,saved=false;
  const source={...record,handoffStatus:'failed',errorMessage:'old 502'};
  const context={cleanSheetWhitespace:x=>String(x||'').trim(),excerptFingerprint,findExcerptAcknowledgement,verifyExcerptReadback,
   poetryPleaseApiKey:'test',poetryPleaseApiUrl:'https://receiver.invalid/api',AbortSignal,
   getExcerptHandoffRecords:async()=>({records:[++reads>1&&changed?{...source,excerpt:'Changed'}:source]}),
   buildPoetryPleaseExcerptRecord:r=>r,invalidateQueueSnapshots(){},
   syncWeaverRuntimeDb:async(action,payload)=>{calls.push({action,payload});if(action==='reconcile_excerpt_handoff'){saved=true;return {saved:true};}
    return {records:saved?[{...source,handoffStatus:'sent',poetryPleaseItemId:'EXC-ONE'}]:[]};},
   fetch:async url=>({ok:true,json:async()=>url.includes('weaverImportHealth')?health:body})};
  const inspect=runInNewContext(`(${server.match(/async function inspectExcerptHandoffReceipt\([^]*?\n}\n/)[0]})`,context);
  if(changed) {await assert.rejects(inspect('handoff','operator',true),/approved source changed/);assert.equal(saved,false);}
  else {assert.equal((await inspect('handoff','operator',true)).saved,true);
   assert.equal(calls.find(c=>c.action==='reconcile_excerpt_handoff').payload.approvedBaseline,true);assert.equal(reads,2);}
 }
});
