import test from 'node:test';
import assert from 'node:assert/strict';
import {INVENTORY_ID,prepareEventIntake,validateEventManifest,currentEventReceipt,footageEventAsSet} from '../footage_event_intake.mjs';
const id='individual_video_12345';
const manifest=()=>({sourceSpreadsheetId:INVENTORY_ID,eventName:'Pilot event',omFolderId:'event_folder_123456789',closed:false,titles:[{sourceRow:6000,performer:'Poet',title:'Poem',rights:'Have video release',contentLink:`https://drive.google.com/file/d/${id}/view?resourcekey=example`} ]});
const file=async n=>({id:n,name:'Clip 1.mov',mimeType:'video/quicktime'});
test('accepted intake preserves event, source row and resource key; acknowledgement does not claim playback or closure',async()=>{
 const r=await prepareEventIntake(manifest(),file);assert.equal(r.acceptedCount,1);assert.equal(r.readyForClosureReview,true);
 assert.equal(r.accepted[0].sourceRow,6000);assert.match(r.accepted[0].sourceVideoUrl,/resourcekey=example/);
 assert.equal(r.playbackVerified,false);assert.equal(r.closedAutomatically,false);assert.equal(r.accepted[0].publicationRestricted,true);
 assert.equal(footageEventAsSet(r).inventoryTitles[id].performer,'Poet');
});
test('retries have stable identity and digest; changed metadata is distinguishable',()=>{
 const a=validateEventManifest(manifest()),b=validateEventManifest({...manifest(),eventName:' Pilot   event '});assert.equal(a.eventId,b.eventId);assert.equal(a.digest,b.digest);
 const m=manifest();m.titles[0].title='Other';assert.equal(a.eventId,validateEventManifest(m).eventId);assert.notEqual(a.digest,validateEventManifest(m).digest);
});
test('partial batches expose blocked records and cannot report closure readiness',async()=>{
 const m=manifest();m.titles.push({...m.titles[0],sourceRow:6001,rights:'No video release'});
 const r=await prepareEventIntake(m,file);assert.equal(r.acceptedCount,1);assert.equal(r.blockedCount,1);assert.equal(r.readyForClosureReview,false);
});
test('folders, arbitrary URLs, unreadable media, non-video files, and duplicate files do not count as accepted',async()=>{
 for(const link of ['https://drive.google.com/drive/folders/event_folder_123456789','https://evil.example/file/d/'+id]){
  const m=manifest();m.titles[0].contentLink=link;assert.equal((await prepareEventIntake(m,file)).acceptedCount,0);
 }
 assert.equal((await prepareEventIntake(manifest(),async()=>{throw Error('403');})).acceptedCount,0);
 assert.equal((await prepareEventIntake(manifest(),async n=>({id:n,mimeType:'application/pdf'}))).acceptedCount,0);
 const m=manifest();m.titles.push({...m.titles[0],sourceRow:6001});assert.equal((await prepareEventIntake(m,file)).blocked[0].reason,'duplicate_video_in_event');
});
test('receipt recheck detects loss of Weaver access',async()=>{
 const r=await prepareEventIntake(manifest(),file);
 const updated=await currentEventReceipt(r,async()=>{throw Error('404');});assert.equal(updated.readyForClosureReview,false);assert.deepEqual(updated.currentUnavailableFileIds,[id]);
});
test('closed events, wrong inventory, no titles and duplicate source rows are rejected',()=>{
 assert.throws(()=>validateEventManifest({...manifest(),closed:true}));assert.throws(()=>validateEventManifest({...manifest(),sourceSpreadsheetId:'other'}));
 assert.throws(()=>validateEventManifest({...manifest(),titles:[]}));const m=manifest();m.titles.push({...m.titles[0]});assert.throws(()=>validateEventManifest(m));
});
