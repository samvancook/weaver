import test from 'node:test';import assert from 'node:assert/strict';
import {makeVideoAssetLinks,planVideoMediaResolution,assertResolvedVideoMedia,withMediaResolution} from '../video_resolution.mjs';
import {createVideoResolutionService} from '../video_resolution_service.mjs';
const c={prioritySetId:'event',sourceFileId:'source',sourceFolderId:'folder',sourceMediaType:'OM',author:'Poet',poemTitle:'Poem',sourceEvent:'Event',reviewCount:5,baseScore:8.5,excerptRecordIds:['e'],gate:{decision:'send_to_editing',publishableAssetUrl:'',selectedExcerptRecordIds:['selected']}};
const fv=(id,w=2160,h=3840)=>({fileId:id,name:'[Vertical Version] Poet - Poem.mov',url:`https://drive.google.com/file/d/${id}/view`,mediaType:'FV',width:w,height:h,byteCount:100,durationMs:10000,checksum:'one'});
const scope={type:'event_scope',prioritySetId:'event',sourceFolderId:'folder',confirmedCoverage:true,finishedFolderIds:['finished'],revision:1};
test('linking and pairing do not mutate release decisions or selected excerpts',()=>{
 const before=structuredClone(c);const r=makeVideoAssetLinks(c,null,{primaryFileId:'a',sourceRole:'OM',confirmPairing:false},'admin@example.com',[fv('a')]);
 assert.equal(r.status,'needs_pairing');assert.deepEqual(c,before);assert.equal(r.decision,undefined);assert.equal(r.selectedExcerptRecordIds,undefined);
 const confirmed=makeVideoAssetLinks(c,r,{primaryFileId:'a',sourceRole:'OM',confirmPairing:true,evidence:'Compared exact recording and finished export'},'admin@example.com',[fv('a')]);assert.equal(confirmed.status,'resolved');assert.deepEqual(c,before);
});
test('unknown, missing scope, failed lookup and duplicate source recordings never request editing',()=>{
 for(const [candidate,s,lookup,status]of [[{...c,sourceMediaType:'unknown'},scope,{complete:true},'needs_classification'],[c,null,{complete:true},'needs_scope'],[c,scope,{complete:false},'lookup_failed'],[c,scope,{complete:true,error:'403'},'lookup_failed'],[c,scope,{complete:true,sourceCount:2},'needs_pairing']]){
  const r=planVideoMediaResolution(candidate,null,s,lookup);assert.equal(r.status,status);assert.equal(r.editingRequest,undefined);
 }
});
test('only confirmed OM and a complete bounded negative lookup produce one stable editing request',()=>{
 const a=planVideoMediaResolution(c,null,scope,{complete:true,files:[]},'2026-11-13T18:00:00Z');assert.equal(a.status,'editing_requested');assert.equal(a.editingRequest.status,'open');
 const b=planVideoMediaResolution(c,a,scope,{complete:true,files:[]},'2026-11-13T19:00:00Z');assert.deepEqual(a.editingRequest,b.editingRequest);
 const done=planVideoMediaResolution(c,b,scope,{complete:true,files:[fv('new')]});assert.equal(done.status,'resolved');assert.equal(done.editingRequest.status,'resolved');assert.equal(done.editingRequest.requestId,a.editingRequest.requestId);
});
test('exact event-scoped matches retain versions and select one unique preferred export',()=>{
 const assets=[fv('v'),fv('small',720,1280),{...fv('h',3840,2160),name:'Poet - Poem.mov'}];
 const r=planVideoMediaResolution(c,null,scope,{complete:true,files:assets});assert.equal(r.status,'resolved');assert.equal(r.assets.length,3);assert.equal(r.primaryFileId,'v');
 assert.equal(planVideoMediaResolution(c,null,scope,{complete:true,files:[fv('one'),fv('two')]}).status,'needs_pairing');
 assert.equal(planVideoMediaResolution(c,null,scope,{complete:true,files:[{...fv('wrong'),name:'Other Poet - Poem.mov'}]}).status,'editing_requested');
});
test('a linked FV that becomes unavailable never produces a replacement editing request',()=>{
 const prior={assets:[fv('a')],primaryFileId:'a',pairing:{status:'verified'}};
 assert.equal(planVideoMediaResolution(c,prior,scope,{error:'File unavailable'}).status,'lookup_failed');
 const r=planVideoMediaResolution(c,prior,scope,{complete:true,files:[]});assert.equal(r.status,'needs_pairing');assert.equal(r.editingRequest,undefined);
});
test('content/version changes revoke pair readiness; unchanged primary does not require repeat confirmation',()=>{
 const prior={sourceRole:'OM',assets:[fv('a')],primaryFileId:'a',pairing:{status:'verified'}};
 assert.equal(planVideoMediaResolution(c,prior,scope,{selectedAsset:{...fv('a'),checksum:'changed'}}).status,'needs_pairing');
 assert.equal(planVideoMediaResolution(c,prior,scope,{selectedAsset:fv('a')}).status,'resolved');
 const linked=makeVideoAssetLinks(c,prior,{primaryFileId:'a',sourceRole:'OM'},'admin@example.com',[fv('a'),fv('b',720,1280)]);assert.equal(linked.status,'resolved');
});
test('direct FV needs exact review lineage; uncertain replacements remain identity checks',()=>{
 const candidate={...c,sourceMediaType:'FV',reviewEvidence:[{mapping:'direct_file_id',sourceFileId:'source'}]};
 assert.equal(planVideoMediaResolution(candidate,null,null,{sourceAsset:fv('source')}).status,'resolved');
 assert.equal(planVideoMediaResolution({...candidate,reviewEvidence:[{mapping:'filename_reconciliation',sourceFileId:'old'}]},null,null,{sourceAsset:fv('source')}).status,'needs_identity');
});
test('delivery requires a fresh verified selected FV matching the release gate',()=>{
 const record=planVideoMediaResolution(c,null,scope,{complete:true,files:[fv('a')]});const ready=withMediaResolution({...c,gate:{...c.gate,publishableAssetUrl:fv('a').url}},record);
 assert.equal(assertResolvedVideoMedia(ready).fileId,'a');
 assert.throws(()=>assertResolvedVideoMedia({...ready,gate:{publishableAssetUrl:fv('other').url}}),/differs/);
 assert.throws(()=>assertResolvedVideoMedia({...ready,mediaResolution:{...ready.mediaResolution,checkedAt:'2000-01-01'}}),/fresh/);
});
test('asset service writes only the resolution resource and wakeup, never a release gate',async()=>{
 const calls=[];const service=createVideoResolutionService({db:async(action,p)=>{calls.push(action);if(action==='get_video_media_resolutions')return{ok:true,records:[]};if(action==='save_video_media_resolution')return{ok:true,record:p.record};throw Error('Forbidden writer '+action);},getCandidates:async()=>({candidates:[c]}),getFile:async id=>({id,name:'Poet - Poem.mov',mimeType:'video/quicktime',size:'100',videoMediaMetadata:{width:2160,height:3840}}),fileRole:async()=>'FV',enqueue:async()=>{calls.push('enqueue');}});
 const r=await service.saveLinks({prioritySetId:'event',sourceFileId:'source',urls:['https://drive.google.com/file/d/a/view'],primaryFileId:'a'},'admin@example.com');assert.equal(r.releaseDecisionChanged,false);assert.deepEqual(calls,['get_video_media_resolutions','save_video_media_resolution','enqueue']);assert.equal(c.gate.decision,'send_to_editing');
});

test('the sender verifies the exact selected file version even when parent metadata is hidden',async()=>{
 const {assertSelectedAssetVersion}=await import('../video_resolution.mjs');const a=fv('a');
 const f={id:'a',mimeType:'video/quicktime',size:100,md5Checksum:'one',videoMediaMetadata:{width:2160,height:3840,durationMillis:'10000'}};
 assert.equal(assertSelectedAssetVersion(a,f),a);assert.throws(()=>assertSelectedAssetVersion(a,{...f,id:'other'}),/unavailable/);assert.throws(()=>assertSelectedAssetVersion(a,{...f,md5Checksum:'changed'}),/checksum/);assert.throws(()=>assertSelectedAssetVersion(a,{...f,videoMediaMetadata:{...f.videoMediaMetadata,height:1280}}),/dimensions/);
});

test('Drive sharing query parameters do not change the selected file identity',()=>{
 const record=planVideoMediaResolution(c,null,scope,{complete:true,files:[fv('a')]});const ready=withMediaResolution({...c,gate:{publishableAssetUrl:fv('a').url+'?usp=drivesdk'}},record);assert.equal(assertResolvedVideoMedia(ready).fileId,'a');assert.throws(()=>assertResolvedVideoMedia({...ready,gate:{publishableAssetUrl:'https://other.example/file/d/a/view'}}),/differs/);
});

test('scheduled resolution moves one durable editing request to verified media when the edit arrives',async()=>{
 const records=[scope];let outputs=[],gate=structuredClone(c.gate);const events=[];
 const db=async(action,p)=>{
  if(action==='get_video_media_resolutions')return{ok:true,records:structuredClone(records)};
  if(action==='save_video_media_resolution'){const i=records.findIndex(r=>r.type==='recording'&&r.sourceFileId===p.record.sourceFileId);const prior=i<0?null:records[i];assert.equal(p.expectedRevision,prior?.revision||0);const saved={...p.record,revision:(prior?.revision||0)+1};if(i<0)records.push(saved);else records[i]=saved;return{ok:true,record:saved};}
  if(action==='prepare_resolved_video_gate'){assert.equal(p.candidate.mediaResolution.status,'resolved');gate={...gate,decision:'ready_for_poetry_please',publishableAssetUrl:p.candidate.mediaResolution.primaryUrl,sourceReviewFileId:'source',videoRecordId:'weaver:video:source'};events.push('policy_ready');return{ok:true,record:gate};}
  throw Error(action);
 };
 const service=createVideoResolutionService({db,getSet:()=>({folderId:'folder'}),getFile:async id=>({id,name:id==='source'?'MVI_1000.MP4 Poet - Poem':'Poet - Poem.mov',mimeType:'video/mp4',size:'100',videoMediaMetadata:{width:2160,height:3840,durationMillis:'10000'}}),fileRole:async url=>url.includes('/source/')?'OM':'FV',folderRole:async()=>({sourceMediaType:'FV'}),listEntries:async()=>outputs});
 const candidate={...c,mediaResolutionRequired:true,videoRecordId:'weaver:video:source'};
 let run=await service.prepare([candidate]);assert.equal(run.candidates[0].mediaResolution.status,'editing_requested');const request=records.find(r=>r.type==='recording').editingRequest.requestId;
 run=await service.prepare([candidate]);assert.equal(records.filter(r=>r.type==='recording').length,1);assert.equal(records.find(r=>r.type==='recording').editingRequest.requestId,request);
 outputs=[{id:'export',name:'[Vertical Version] Poet - Poem.mov',mimeType:'video/mp4',size:'100',videoMediaMetadata:{width:2160,height:3840,durationMillis:'10000'}}];
 run=await service.prepare([candidate]);assert.equal(run.candidates[0].mediaResolution.status,'resolved');assert.equal(run.candidates[0].gate.decision,'ready_for_poetry_please');assert.equal(records.find(r=>r.type==='recording').editingRequest.status,'resolved');assert.deepEqual(events,['policy_ready']);
});

test('confirmed alternate review files share the known recording, while unconfirmed references cannot alias it',async()=>{
 const root={...c,sourceReviewFileId:'source',videoRecordId:'weaver:video:source'},other={...c,sourceFileId:'other',sourceReviewFileId:'',sourceMediaType:'FV'};
 const record={type:'recording',prioritySetId:'event',sourceFileId:'source',reviewedSourceFileId:'source',sourceRole:'OM',status:'resolved',pairing:{status:'verified'},assets:[{...fv('other'),pairingVerified:true}],primaryFileId:'other'};
 const service=createVideoResolutionService({db:async()=>({ok:true,records:[record]}),getSet:()=>({folderId:'folder'})});
 const views=await service.attach([root,other]);assert.equal(views[1].sourceReviewFileId,'source');assert.equal(views[1].videoRecordId,'weaver:video:source');assert.equal(views[1].mediaResolution.primaryFileId,'other');
 record.assets[0].pairingVerified=false;const unresolved=await service.attach([other]);assert.equal(unresolved[0].sourceReviewFileId,'');
});

test('finding an old FV cannot override explicit human rework instructions',()=>{
 const candidate={...c,gate:{...c.gate,decidedBy:'editor@example.com',decidedAt:'2026-11-13T18:00:00Z',editingInstructions:'Remove incorrect title card'}};
 assert.equal(planVideoMediaResolution(candidate,null,scope,{complete:true,files:[fv('old')]}).status,'needs_pairing');
 const prior={assets:[fv('new')],primaryFileId:'new',pairing:{status:'verified',method:'administrator_confirmation',verifiedAt:'2026-11-13T19:00:00Z'}};
 assert.equal(planVideoMediaResolution(candidate,prior,scope,{selectedAsset:fv('new')}).status,'resolved');
});
