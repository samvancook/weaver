import test from 'node:test';import assert from 'node:assert/strict';
import {INVENTORY_ID,inventoryCurationRights,prepareEventIntake,currentEventReceipt,footageEventAsSet,videoPublicationRestriction} from '../footage_event_intake.mjs';
import {resolveEventTarget} from '../footage_event_links.mjs';
const title=(sourceRow,rights,id='individual_video_'+sourceRow)=>({sourceRow,performer:'Poet',title:'Poem '+sourceRow,rights,contentLink:`https://drive.google.com/file/d/${id}/view`});
const manifest=titles=>({sourceSpreadsheetId:INVENTORY_ID,eventName:'Rights test',omFolderId:'event_folder_12345678',closed:false,titles});
const file=async id=>({id,mimeType:'video/mp4',name:'Video '+id});
test('agreed labels normalize whitespace/case without changing unknown-label handling',()=>{
 assert.equal(inventoryCurationRights(' Title   not included on Video Release '),'include');
 assert.equal(inventoryCurationRights('Performance Opt Out'),'exclude');
 assert.equal(inventoryCurationRights('No video release'),'exclude');
 assert.equal(inventoryCurationRights('some new value'),'review');
});
test('mixed rights preserve source identities, exclude without fetching, and keep publication/closure separate',async()=>{
 const input=manifest([title(2,'Have video release'),title(3,'Title not included on Video Release'),title(4,'No video release'),title(5,'Performance Opt Out')]);
 input.titles[2].contentLink='not a file';input.titles[3].contentLink='';const before=JSON.stringify(input),calls=[];
 const r=await prepareEventIntake(input,async id=>{calls.push(id);return file(id);});
 assert.equal(JSON.stringify(input),before);assert.equal(r.submittedCount,4);assert.equal(r.eligibleCount,2);assert.equal(r.acceptedCount,2);assert.equal(r.excludedCount,2);assert.equal(r.blockedCount,0);assert.equal(r.transferComplete,true);
 assert.deepEqual(calls,['individual_video_2','individual_video_3']);assert.deepEqual(r.accepted.map(t=>t.sourceRow),[2,3]);assert.deepEqual(r.excluded.map(t=>t.sourceRow),[4,5]);assert.ok(r.excluded.every(t=>t.reason==='excluded_from_curation_by_inventory_rights'));
 assert.ok(r.accepted.every(t=>t.publicationRestricted));assert.equal(r.closedAutomatically,false);assert.equal(r.readyForClosureReview,false);
 const set=footageEventAsSet(r);assert.equal(set.inventoryFiles.length,2);assert.equal(videoPublicationRestriction('Poem',set,'individual_video_3'),'inventory_rights_unverified');
 const reread=await currentEventReceipt(r,file);assert.equal(reread.excludedCount,2);assert.equal(reread.transferComplete,true);
});
test('all-excluded event never reports transfer complete and never fetches media',async()=>{
 const r=await prepareEventIntake(manifest([title(2,'No video release'),title(3,'Performance Opt Out')]),()=>{throw Error('Must not fetch');});
 assert.equal(r.acceptedCount,0);assert.equal(r.eligibleCount,0);assert.equal(r.excludedCount,2);assert.equal(r.blockedCount,0);assert.equal(r.transferComplete,false);
});
test('unknown rights, bad eligible links and unreadable eligible media still block',async()=>{
 const m=manifest([title(2,'unknown'),title(3,'Title not included on Video Release')]);
 const r=await prepareEventIntake(m,()=>{throw Error('403');});assert.deepEqual(r.blocked.map(t=>t.reason),['rights_need_review','weaver_cannot_read_video']);assert.equal(r.transferComplete,false);
 m.titles[1].contentLink='https://drive.google.com/drive/folders/event_folder_12345678';assert.equal((await prepareEventIntake(m,file)).blocked[1].reason,'individual_drive_video_required');
});
test('excluded assets do not force a new target, collide with another event or demand missing correspondences',async()=>{
 const m=manifest([title(2,'Have video release'),title(3,'No video release')]);const sets=new Map([['existing',{id:'existing',label:'Existing'}],['other',{id:'other',inventoryEvent:true}]]);
 const target=await resolveEventTarget(m,sets,async s=>[{id:s.id==='existing'?'individual_video_2':'individual_video_3'}],u=>u.match(/\/d\/([^/]+)/)?.[1]||'');assert.equal(target.mode,'existing_event');assert.equal(target.prioritySetId,'existing');
 const r=await prepareEventIntake(m,file,target);assert.equal(r.transferComplete,true);assert.equal(r.accepted[0].sourceFileId,'individual_video_2');assert.equal(r.excludedCount,1);assert.equal(footageEventAsSet(r),null);
});
