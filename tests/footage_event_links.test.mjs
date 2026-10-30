import test from 'node:test';
import assert from 'node:assert/strict';
import {INVENTORY_ID,prepareEventIntake,currentEventReceipt,footageEventAsSet,fileId} from '../footage_event_intake.mjs';
import {REVIEWED_EVENT_LINKS,resolveEventTarget} from '../footage_event_links.mjs';
const link=REVIEWED_EVENT_LINKS[0];
const contract='All Button Authors, All features, All workshop instructors - contract is not stored with video releases';
const input=()=>({sourceSpreadsheetId:INVENTORY_ID,eventName:'PAGES footage from Rudy Francisco 2026',omFolderId:link.inventoryFolderId,closed:false,
 titles:link.pairs.map(([id],i)=>({sourceRow:6871+i,performer:'Pages Matam',title:['birth chart','Lazarus','wedding','Accent'][i],rights:contract,contentLink:`https://drive.google.com/file/d/${id}/view`}))});
const set={id:link.prioritySetId,label:'PAGES Matam in San Diego - 2026'};
const sets=()=>new Map([[set.id,set]]);
const list=async()=>link.pairs.map(([,id])=>({id}));
const getFile=async id=>({id,mimeType:'video/mp4',name:id,videoMediaMetadata:link.pairs.some(([x])=>x===id)?{width:3840,height:2160,durationMillis:'200000'}:{width:2160,height:3840,durationMillis:'190000'}});
test('PAGES reuses its existing queue and stores distinct HV/VV assets without inventing canonical text identity',async()=>{
 const target=await resolveEventTarget(input(),sets(),list,fileId);
 const r=await prepareEventIntake(input(),getFile,target);
 assert.equal(r.mode,'existing_event');assert.equal(r.prioritySetId,set.id);assert.equal(r.acceptedCount,4);assert.equal(r.blockedCount,0);assert.equal(r.transferComplete,true);
 assert.equal(r.readyForClosureReview,false);assert.equal(r.eventCompleteness,'not_assessed');assert.equal(footageEventAsSet(r),null);
 r.accepted.forEach((t,i)=>{
  assert.equal(t.inventorySourceFileId,link.pairs[i][0]);assert.equal(t.sourceFileId,link.pairs[i][1]);assert.notEqual(t.inventorySourceFileId,t.sourceFileId);
  assert.deepEqual(t.assets.map(a=>a.format),['HV','VV']);assert.equal(t.textReference.canonicalTextId,null);
  assert.deepEqual(t.assets.map(a=>a.mediaType),['OM','FV']);
  assert.equal(t.workflow.originalMediaFileId,link.pairs[i][0]);
  assert.equal(t.workflow.finishedVideoFileId,link.pairs[i][1]);
  assert.equal(t.workflow.publicationAssetFileId,link.pairs[i][1]);
  assert.equal(t.workflow.editingStatus,'finished_video_linked');
  assert.equal(t.workflow.publicationStatus,'awaiting_curation_approval');
  assert.equal(t.exactDerivationVerified,false);assert.equal(t.publicationRestricted,true);
 });
 const retry=await prepareEventIntake(input(),getFile,target);
 assert.equal(retry.eventId,r.eventId);assert.equal(retry.routingDigest,r.routingDigest);assert.deepEqual(retry.accepted.map(t=>t.familyId),r.accepted.map(t=>t.familyId));
});
test('media role is independent of orientation, and unreviewed files remain unclassified',async()=>{
 const target=await resolveEventTarget(input(),sets(),list,fileId);
 const r=await prepareEventIntake(input(),async id=>({...await getFile(id),videoMediaMetadata:{width:3840,height:2160}}),target);
 assert.deepEqual(r.accepted[0].assets.map(a=>a.format),['HV','HV']);
 assert.deepEqual(r.accepted[0].assets.map(a=>a.mediaType),['OM','FV']);
 const unknown=await prepareEventIntake({...input(),omFolderId:'unreviewed_folder_12345'},getFile);
 assert.equal(unknown.accepted[0].assets[0].mediaType,'unknown');
 assert.equal(unknown.accepted[0].workflow.publicationAssetFileId,null);
 assert.equal(unknown.accepted[0].workflow.editingStatus,'media_type_needs_review');
 const om=await prepareEventIntake(input(),getFile,{mode:'new_event',inventoryMediaType:'OM',mediaTypeEvidence:'Reviewed OM source'});
 assert.equal(om.accepted[0].workflow.editingStatus,'needs_editing');
 assert.equal(om.accepted[0].workflow.publicationAssetFileId,null);
 assert.equal(om.accepted[0].workflow.publicationStatus,'awaiting_finished_video');
 assert.equal(footageEventAsSet(om).sourceMediaType,'OM');
 assert.equal(footageEventAsSet(unknown).sourceMediaType,'unknown');
});
test('a missing target queue fails instead of silently creating a duplicate',async()=>{
 await assert.rejects(resolveEventTarget(input(),new Map(),list,fileId),/unavailable/);
 const target=await resolveEventTarget(input(),sets(),async()=>[],fileId);
 const r=await prepareEventIntake(input(),getFile,target);assert.equal(r.acceptedCount,0);assert.equal(r.blockedCount,4);assert.equal(r.transferComplete,false);
});
test('unmapped additional inventory assets are explicit and never guessed from a matching title',async()=>{
 const m=input();m.titles.push({...m.titles[0],sourceRow:7000,contentLink:'https://drive.google.com/file/d/unmapped_individual_123/view'});
 const r=await prepareEventIntake(m,getFile,await resolveEventTarget(m,sets(),list,fileId));
 assert.equal(r.acceptedCount,4);assert.equal(r.blocked[0].reason,'existing_event_asset_connection_missing');assert.equal(r.transferComplete,false);
});
test('current receipt checks both versions and existing queue membership',async()=>{
 const r=await prepareEventIntake(input(),getFile,await resolveEventTarget(input(),sets(),list,fileId));
 const failed=await currentEventReceipt(r,async id=>{if(id===link.pairs[0][0])throw Error('403');return getFile(id);});
 assert.equal(failed.transferComplete,false);assert.deepEqual(failed.currentUnavailableFileIds,[link.pairs[0][0]]);
 assert.equal((await currentEventReceipt(r,getFile,async()=>false)).transferComplete,false);
});
test('generic exact-ID matches reuse one existing queue; partial or ambiguous matches cannot create new queues',async()=>{
 const m={...input(),omFolderId:'other_folder_123456789',titles:[input().titles[0]]};
 const file={id:link.pairs[0][0]};
 const a={id:'existing-a',label:'A'},b={id:'existing-b',label:'B'};
 const exact=await resolveEventTarget(m,new Map([[a.id,a]]),async()=>[file],fileId,[]);
 assert.equal(exact.prioritySetId,a.id);assert.equal(exact.pairs.get(file.id),file.id);
 await assert.rejects(resolveEventTarget(m,new Map([[a.id,a],[b.id,b]]),async()=>[file],fileId,[]),/multiple queues/);
 const mixed={...m,titles:[...m.titles,{...m.titles[0],sourceRow:7000,contentLink:'https://drive.google.com/file/d/new_individual_123456/view'}]};
 await assert.rejects(resolveEventTarget(mixed,new Map([[a.id,a]]),async()=>[file],fileId,[]),/Some videos/);
 const fresh=await resolveEventTarget(m,new Map([[a.id,a]]),async()=>[],fileId,[]);assert.equal(fresh.mode,'new_event');
});
test('retrying a previously created inventory queue does not reinterpret it as another existing event',async()=>{
 const {validateEventManifest}=await import('../footage_event_intake.mjs');
 const m={...input(),omFolderId:'fresh_folder_123456789'};
 const ownId=validateEventManifest(m).eventId;
 const own={id:ownId,inventoryEvent:true};
 const target=await resolveEventTarget(m,new Map([[ownId,own]]),async()=>m.titles.map(t=>({id:fileId(t.contentLink)})),fileId,[]);
 assert.equal(target.mode,'new_event');
 await assert.rejects(resolveEventTarget(m,new Map([['other',{id:'other',inventoryEvent:true}]]),async()=>m.titles.map(t=>({id:fileId(t.contentLink)})),fileId,[]),/another inventory event/);
});
test('multiple distinct inventory versions can share one reviewed asset family without multiplying the curation queue',async()=>{
 const m=input();m.titles=m.titles.slice(0,2);
 const target={mode:'existing_event',prioritySetId:set.id,label:set.label,
  pairs:new Map([[link.pairs[0][0],link.pairs[0][1]],[link.pairs[1][0],link.pairs[0][1]]]),available:new Set([link.pairs[0][1]])};
 const r=await prepareEventIntake(m,getFile,target);
 assert.equal(r.acceptedCount,2);assert.equal(r.blockedCount,0);assert.equal(r.inventoryAssetCount,2);assert.equal(r.curationAssetCount,1);assert.equal(r.familyCount,1);assert.equal(footageEventAsSet(r),null);
});
