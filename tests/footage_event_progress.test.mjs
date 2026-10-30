import test from 'node:test';
import assert from 'node:assert/strict';
import {buildFootageProgress,footageProvenanceForFile} from '../footage_event_progress.mjs';
import {currentEventReceipt} from '../footage_event_intake.mjs';
const record={prioritySetId:'event-a',acceptedCount:1,blockedCount:0,accepted:[{sourceFileId:'fv-a',assets:[{driveFileId:'om-a'},{driveFileId:'fv-a'}]}]};
const review=(extra={})=>({prioritySetId:'event-a',sourceFileId:'fv-a',reviewerEmail:'a@example.com',rating:'like',updatedAt:'2026-10-30T10:00:00Z',excerptRecordIds:['excerpt-a'],...extra});
const gate=(extra={})=>({prioritySetId:'event-a',sourceFileId:'fv-a',decision:'ready_for_poetry_please',...extra});
test('provenance retains exact inventory/OM/FV identities without borrowing another title or text ID',()=>{
 const connections=[{sourceFileId:'fv-a',inventoryEventId:'inventory-a',inventoryEventName:'Event A',sourceSpreadsheetId:'sheet-a',sourceRow:6871,familyId:'family-a',workflow:{originalMediaFileId:'om-a',finishedVideoFileId:'fv-a'}},{sourceFileId:'other',sourceRow:6872}];
 assert.deepEqual(footageProvenanceForFile(connections,'fv-a'),[{inventoryEventId:'inventory-a',inventoryEventName:'Event A',sourceSpreadsheetId:'sheet-a',sourceRow:6871,familyId:'family-a',originalMediaFileId:'om-a',finishedVideoFileId:'fv-a',curationFileId:'fv-a'}]);
});
test('progress counts the latest review once per person and only the exact event/file',()=>{
 const r=buildFootageProgress(record,[review(),review({reviewerEmail:'A@EXAMPLE.COM',updatedAt:'2026-10-30T11:00:00Z',excerptRecordIds:['excerpt-b']}),review({prioritySetId:'other'}),review({sourceFileId:'om-a'})],[],'now').byFileId['fv-a'];
 assert.equal(r.reviewCount,1);assert.equal(r.linkedExcerptCount,1);assert.equal(r.state,'curation_in_progress');
 assert.equal(r.poetryPlease.status,'not_sent');assert.equal(r.calendarStatus,'not_checked');
});
test('a saved PP receipt requires the exact reviewed source identity and cannot claim calendar placement',()=>{
 const receipt={ok:true,status:'sent_to_poetry_please',sourceRecordId:'weaver:video:fv-a',canonicalVideoId:'PP-A',canonicalVideoUrl:'https://poetryplease.org/?item=PP-A',durableCopyStatus:'pending'};
 const p=buildFootageProgress(record,[],[gate({poetryPleaseHandoff:receipt})]).byFileId['fv-a'];
 assert.equal(p.state,'received_by_poetry_please');assert.equal(p.poetryPlease.canonicalVideoId,'PP-A');assert.equal(p.poetryPlease.durableCopyStatus,'pending');assert.equal(p.calendarStatus,'not_checked');
 const bad=buildFootageProgress(record,[],[gate({poetryPleaseHandoff:{...receipt,sourceRecordId:'weaver:video:another'}})]).byFileId['fv-a'];
 assert.equal(bad.state,'handoff_needs_attention');assert.equal(bad.poetryPlease.status,'receipt_identity_unverified');assert.equal(bad.poetryPlease.canonicalVideoId,null);
});
test('editing, held decisions and failed or running handoffs remain distinct',()=>{
 for(const [g,state] of [[gate({decision:'send_to_editing'}),'editing_requested'],[gate({decision:'hold'}),'on_hold'],[gate({poetryPleaseHandoff:{status:'failed',error:'receiver unavailable'}}),'handoff_needs_attention'],[gate({poetryPleaseHandoff:{status:'sending'}}),'handoff_in_progress']]) {
  assert.equal(buildFootageProgress(record,[],[g]).byFileId['fv-a'].state,state);
 }
});
test('an unavailable progress source does not masquerade as zero reviews or remove saved links',async()=>{
 const r=await currentEventReceipt(record,async id=>({id,mimeType:'video/mp4'}),async()=>true,async()=>{throw Error('ledger unavailable');});
 assert.equal(r.progressCheck.status,'unavailable');assert.equal(r.accepted[0].progress,null);assert.equal(r.transferComplete,true);assert.equal(r.accepted[0].assets.length,2);
});
