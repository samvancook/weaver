import test from 'node:test';import assert from 'node:assert/strict';
import {refreshAcceptedVideoReceipt} from '../video_receipt_refresh.mjs';
const gate={prioritySetId:'event',sourceFileId:'accepted-original-file',poetryPleaseHandoff:{status:'sent_to_poetry_please',sourceRecordId:'weaver:video:root',canonicalVideoId:'canonical'}};
const candidate={prioritySetId:'event',sourceFileId:'new-finished-representative',videoRecordId:'weaver:video:root',gate};
const receipt={sourceRecordId:'weaver:video:root',canonicalVideoId:'canonical',durableCopyStatus:'verified'};
const parseReceipt=(r,g)=>{assert.equal(r.sourceRecordId,g.poetryPleaseHandoff.sourceRecordId);assert.equal(r.canonicalVideoId,g.poetryPleaseHandoff.canonicalVideoId);return r;};
test('grouped performance refreshes the original accepted gate, preserving legacy empty attempt IDs',async()=>{
 let calls=0;const original=structuredClone(candidate);
 const result=await refreshAcceptedVideoReceipt({candidate,receipt,parseReceipt,db:async(action,payload)=>{calls++;assert.equal(action,'refresh_video_curation_receipt');assert.equal(payload.sourceFileId,'accepted-original-file');assert.equal(payload.attemptId,'');assert.equal(payload.receipt.canonicalVideoId,'canonical');return{saved:true};}});
 assert.equal(result.saved,true);assert.equal(calls,1);assert.deepEqual(candidate,original);
});
test('unchanged receipt has no write; gate and receiver identity mismatches cannot write',async()=>{
 const db=async()=>{throw Error('must not write');};
 assert.equal((await refreshAcceptedVideoReceipt({candidate:{...candidate,gate:{...gate,poetryPleaseHandoff:{...gate.poetryPleaseHandoff,durableCopyStatus:'verified'}}},receipt,parseReceipt,db})).unchanged,true);
 for(const patch of [{prioritySetId:'other'},{videoRecordId:'weaver:video:wrong'}])await assert.rejects(refreshAcceptedVideoReceipt({candidate:{...candidate,...patch},receipt,parseReceipt,db}),/does not match/);
 await assert.rejects(refreshAcceptedVideoReceipt({candidate,receipt:{...receipt,canonicalVideoId:'wrong'},parseReceipt,db}));
});
test('concurrent/changed handoff stays unresolved with the actual reason',async()=>{
 await assert.rejects(refreshAcceptedVideoReceipt({candidate,receipt,parseReceipt,db:async()=>({saved:false,reason:'handoff_attempt_changed'})}),/handoff_attempt_changed/);
});
