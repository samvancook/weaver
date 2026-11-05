import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import {readFileSync} from 'node:fs';
import {createHash} from 'node:crypto';
import {assertVideoRetryPreflight,assertVideoImportPreflight,assertVideoClaimMatchesRecord,assertVideoImportReadback,driveFileId} from '../video_curation_handoff.mjs';

const source=readFileSync(new URL('../server.mjs',import.meta.url),'utf8');
const functionSource=source.match(/async function handoffVideoCurationGate\([^]*?^}/m)[0];
const record={sourceRecordId:'weaver:video:original',sourceEvent:'event',sourceEventLabel:'Event',finalAssetUrl:'https://drive.google.com/file/d/final/view',reviews:[{reviewId:'r1'}],selectedExcerptRecordIds:['e1']};
const canonical=`WEAVER-VV-${createHash('sha256').update(record.sourceRecordId).digest('hex').slice(0,24).toUpperCase()}`;
const gate={decision:'ready_for_poetry_please',videoRecordId:record.sourceRecordId,sourceEvent:'event',sourceEventLabel:'Event',publishableAssetUrl:record.finalAssetUrl,selectedExcerptRecordIds:['e1'],poetryPleaseHandoff:{status:'failed',sourceRecordId:record.sourceRecordId,error:'timeout'}};
const empty={ok:true,received:false,sourceRecordId:record.sourceRecordId,canonicalVideoId:'',expectedCanonicalVideoId:canonical,review:{reviewCount:0,selectedExcerptIdCount:0},ledger:{count:1,latest:{status:'failed',error:'timeout',importLedgerId:'old-ledger',canonicalVideoId:'',receivedReviewCount:0,receivedSelectedExcerptIdCount:0}}};
function harness(changedBeforeSend=false){
  const calls=[];let receiptReads=0;let claimed;
  const ctx=vm.createContext({AbortSignal,Date,poetryPleaseApiKey:'synthetic',poetryPleaseApiUrl:'https://pp.example/api',cleanSheetWhitespace:x=>String(x||'').trim(),driveFileId,
    getVideoCurationCandidates:async()=>({candidates:[{prioritySetId:'event',sourceFileId:'final',sourceMediaType:'FV',excerptRecordIds:['e1'],gate}]}),
    buildWeaverVideoImport:()=>record,assertVideoRetryPreflight,assertVideoImportPreflight,assertVideoClaimMatchesRecord,assertVideoImportReadback,
    parsePoetryPleaseVideoImport:()=>({ok:true,status:'sent_to_poetry_please',canonicalVideoId:canonical}),
    syncWeaverRuntimeDb:async(action,payload)=>{calls.push(action);if(action==='claim_video_curation_handoff'){assert.equal(payload.retryFailedProof.importLedgerId,'old-ledger');claimed={...gate,poetryPleaseHandoff:{status:'sending',attemptId:'new-attempt',previousHandoff:gate.poetryPleaseHandoff}};return {claimed:true,record:claimed};} if(action==='release_video_curation_handoff_claim')return {released:true};if(action==='finish_video_curation_handoff')return {saved:true,record:{...gate,poetryPleaseHandoff:payload.poetryPleaseHandoff}};throw Error(action);},
    fetch:async(url,options)=>{if(url.includes('Receipt?')){receiptReads++;calls.push('receipt');const body=receiptReads===1?empty:receiptReads===2?(changedBeforeSend?{...empty,received:true,canonicalVideoId:canonical}:empty):{ok:true,received:true,sourceRecordId:record.sourceRecordId,canonicalVideoId:canonical,event:{sourceEvent:'event',sourceEventLabel:'Event'},media:{driveLink:record.finalAssetUrl,reviewOnly:true},review:{reviewCount:1,selectedExcerptIdCount:1}};return {ok:true,json:async()=>body};}calls.push('import');assert.equal(JSON.parse(options.body).sourceRecordId,record.sourceRecordId);return {ok:true,json:async()=>({})};}
  });vm.runInContext(functionSource,ctx);return {ctx,calls};
}

test('retry rechecks receiver after exclusive gate claim, sends once, reads back and finalizes',async()=>{
 const {ctx,calls}=harness();const result=await ctx.handoffVideoCurationGate({prioritySetId:'event',sourceFileId:'final',retryFailed:true});
 assert.equal(result.ok,true);assert.deepEqual(calls,['receipt','claim_video_curation_handoff','receipt','import','receipt','finish_video_curation_handoff']);
});
test('changed receiver before retry restores the unsent claim and never imports',async()=>{
 const {ctx,calls}=harness(true);await assert.rejects(ctx.handoffVideoCurationGate({prioritySetId:'event',sourceFileId:'final',retryFailed:true}));
 assert.deepEqual(calls,['receipt','claim_video_curation_handoff','receipt','release_video_curation_handoff_claim']);
});
