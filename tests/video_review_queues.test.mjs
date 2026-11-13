import {videoFileMediaType} from '../video_media_classification.mjs';
import {VIDEO_AUTO_POLICY,publicVideoPipelineState,videoPipelineKey} from '../video_pipeline.mjs';
import test from 'node:test';import assert from 'node:assert/strict';import fs from 'node:fs';import vm from 'node:vm';
import {videoReviewProgress,videoLibraryStatus,reconcileVideoReviews,resolveVideoReviewSourceFileId,poetryPleaseVideoIdForSourceFile,requiresManualVideoIdentityReview} from '../video_curation_handoff.mjs';
const app=fs.readFileSync(new URL('../public/app.js',import.meta.url),'utf8'),server=fs.readFileSync(new URL('../server.mjs',import.meta.url),'utf8');
const fn=(source,name)=>source.match(new RegExp(`^(?:async )?function ${name}\\([\\s\\S]*?^}`, 'm'))[0];
const reviews=(id,n=5)=>Array.from({length:n},(_,i)=>({sourceFileId:id,prioritySetId:'event',reviewerEmail:`r${i}@example.com`,rating:'moved_me',excerptRecordIds:i===0?['one','two','three']:[]}));
test('distinct reviewers and unique excerpts determine completion, separately for every file',()=>{
 const rated=reviews('a');assert.equal(videoReviewProgress(rated).reviewComplete,true);
 const reconciled=reconcileVideoReviews([{id:'a',name:'a'}],[...rated,...rated.map(r=>({...r,reviewerEmail:r.reviewerEmail.toUpperCase()}))]);
 assert.equal(videoReviewProgress(reconciled).reviewerCount,5);assert.equal(videoReviewProgress(reconciled).excerptCount,3);
 assert.equal(videoReviewProgress(reviews('a',4)).reviewComplete,false);assert.equal(videoReviewProgress(rated.map(r=>({...r,excerptRecordIds:[]}))).reviewComplete,false);
 assert.equal(videoReviewProgress([{...rated[0],ratingSource:'legacy_import',legacyScore:0}]).candidateScore,0);
});
test('collecting excerpts changes progress without raising the review score or crossing the auto threshold',()=>{
 const ratings=['moved_me','moved_me','moved_me','like','like'];
 const original=ratings.map((rating,i)=>({reviewerEmail:`r${i}@example.com`,rating,excerptRecordIds:[]}));
 const collected=original.map((r,i)=>({...r,excerptRecordIds:i<2?['one','two','three']:[]}));
 assert.equal(videoReviewProgress(original).candidateScore,9);
 assert.equal(videoReviewProgress(collected).candidateScore,9);
 assert.equal(videoReviewProgress(collected).excerptCount,3);
 assert.equal(videoReviewProgress(collected).reviewComplete,true);
 const below=collected.map(r=>({...r,ratingSource:'legacy_import',legacyScore:8.49}));
 assert.equal(videoReviewProgress(below).candidateScore,8.49);
 assert.equal(videoLibraryStatus(videoReviewProgress(below),'FV','',null),'review_complete');
 assert.equal(videoLibraryStatus(videoReviewProgress(below),'FV','',{poetryPleaseHandoff:{status:'sent_to_poetry_please'}}),'sent');
 assert.equal(videoReviewProgress(collected.map(r=>({...r,rating:'moved_me'}))).candidateScore,10);
 assert.equal(videoReviewProgress([]).candidateScore,0);
});
test('review completion cannot turn OM, restrictions, holds or uncertain sends into ready-for-handoff',()=>{
 const p=videoReviewProgress(reviews('a'));assert.equal(videoLibraryStatus(p,'FV','',null),'ready_for_handoff');
 assert.equal(videoLibraryStatus(p,'OM','',null),'needs_editing');assert.equal(videoLibraryStatus(p,'FV','unlinked',null),'needs_check');
 assert.equal(videoLibraryStatus(p,'FV','',{decision:'hold'}),'review_complete');assert.equal(videoLibraryStatus(p,'unknown','',null),'needs_check');assert.equal(videoLibraryStatus(p,'FV','',null,false),'needs_check');
 assert.equal(videoLibraryStatus(p,'FV','',{poetryPleaseHandoff:{status:'sending'}}),'needs_check');
 assert.equal(videoLibraryStatus({...p,candidateScore:6},'FV','',null),'review_complete');
 assert.equal(videoLibraryStatus({...p,candidateScore:7},'FV','',{decision:'ready_for_poetry_please'}),'ready_for_handoff');
 assert.equal(videoLibraryStatus({...p,reviewComplete:false},'FV','',{poetryPleaseHandoff:{status:'sent_to_poetry_please'}}),'sent');
});
function serverHarness(){
 const calls=[],data=[...reviews('ready'),...reviews('pending',1).map(r=>({...r,reviewerEmail:'me@example.com'})),...reviews('raw')];
 const gateRecords=[];
 const sets=new Map([['event',{id:'event',label:'Finals',folderId:'folder'}]]),files=['ready','pending','raw','unseen'].map(id=>({id,name:id}));
 const s={videoFileMediaType,VIDEO_AUTO_POLICY,publicVideoPipelineState,videoPipelineKey,PRIORITY_VIDEO_SETS:sets,cleanSheetWhitespace:x=>String(x||'').trim(),videoReviewProgress,videoLibraryStatus,reconcileVideoReviews,resolveVideoReviewSourceFileId,poetryPleaseVideoIdForSourceFile,requiresManualVideoIdentityReview,
 syncWeaverRuntimeDb:async(action)=>{calls.push(action);if(action==='get_curation_reviews')return{ok:true,records:data};if(action==='get_video_curation_gates')return{ok:true,records:gateRecords};if(action==='get_video_pipeline')return{ok:true,records:[]};throw Error('Unexpected writer');},
 getVideoSetMediaType:async()=>({sourceMediaType:'FV'}),listVideoSetFiles:async()=>files,getFootageAssetConnections:async()=>[],isPriorityVideoFileIncluded:()=>true,
 getPriorityVideoReleaseStatus:name=>name==='raw'?'unverified':'',parsePriorityVideoFileName:name=>({author:'Poet',poemTitle:name}),footageProvenanceForFile:()=>[],buildVideoCurationCandidateId:(set,id)=>set+':'+id,encodeURIComponent};
 vm.createContext(s);vm.runInContext(fn(server,'loadVideoReviewLibrary')+'\n'+fn(server,'getPriorityVideoSetAvailability'),s);return{s,calls,data,gateRecords};
}
test('mixed-event read model retains all files, hides completed individuals, and exposes only the requesting reviewer notes',async()=>{
 const {s,calls}=serverHarness(),d=await s.loadVideoReviewLibrary(' ME@example.com ','event');
 assert.equal(d.items.length,4);assert.equal(d.items.find(i=>i.sourceFileId==='ready').needsPriorityReview,false);
 assert.equal(d.items.find(i=>i.sourceFileId==='pending').needsPriorityReview,true);
 assert.equal(d.items.find(i=>i.sourceFileId==='raw').libraryStatus,'needs_check');
 assert.equal(d.sets[0].remainingCount,1);assert.equal(d.sets[0].reviewedCount,1);
 assert.equal(d.items.find(i=>i.sourceFileId==='ready').review,null);assert.equal(d.items.find(i=>i.sourceFileId==='pending').review.rating,'moved_me');
 assert.ok(!JSON.stringify(d).includes('r0@example.com'));assert.ok(calls.every(x=>x.startsWith('get_')));
 const available=await s.getPriorityVideoSetAvailability('me@example.com');assert.equal(available.sets[0].readyVideos,2);assert.equal(available.sets[0].remainingCount,1);
 await assert.rejects(s.loadVideoReviewLibrary('me@example.com','unknown'),/Unknown/);
 s.syncWeaverRuntimeDb=async()=>({ok:false,error:'Unavailable'});await assert.rejects(s.loadVideoReviewLibrary('me@example.com'),/could not be read/);
});
test('personal, reviewed, library, search and status views preserve completed and sent videos for retrieval',()=>{
 const filter=vm.runInNewContext(fn(app,'filterVideoReviewItems')+';filterVideoReviewItems');
 const items=[{author:'Poet',poemTitle:'Pending',needsPriorityReview:true},{author:'Poet',poemTitle:'Done',needsPriorityReview:false,libraryStatus:'review_complete'},
 {author:'Poet',poemTitle:'Mine',review:{rating:'like'},needsPriorityReview:true},{author:'Poet',poemTitle:'Sent',needsPriorityReview:false,libraryStatus:'sent'}];
 assert.deepEqual(Array.from(filter(items,'todo','me').map(i=>i.poemTitle)),['Pending']);
 assert.deepEqual(Array.from(filter(items,'reviewed','me').map(i=>i.poemTitle)),['Mine']);assert.equal(filter(items,'library','').length,4);
 assert.equal(filter(items,'todo','').length,0);assert.equal(filter(items,'library','', 'sent','sent').length,1);
 assert.equal(filter(items,'library','','done')[0].poemTitle,'Done');
});
test('saving removes a new review from todo, leaves it in history, and uses the item event in an all-event library',()=>{
 const item={prioritySetId:'event',sourceFileId:'one',author:'Poet',poemTitle:'Poem',needsPriorityReview:true};
 const ctx={contributorShellParams:new URLSearchParams(),elements:{gatheringVideoView:{value:'todo'},gatheringEmail:{value:'me@example.com',checkValidity:()=>true},gatheringVideoScore:{value:'like'}},
 currentGatheringVideoPlaylist:{allItems:[item],items:[item],reviewerEmail:'me@example.com',index:0},applyGatheringVideoPlaylistItem:()=>{},clearGatheringVideoPlaylistItemFields:()=>{},renderVideoLibraryResults:()=>{},updateGatheringVideoPlaylistUi:()=>{},cleanSheetWhitespace:x=>String(x||'').trim()};
 vm.createContext(ctx);vm.runInContext(['videoReviewView','filterVideoReviewItems','filterGatheringVideoPlaylist','getCurrentGatheringVideoPlaylistItem','buildGatheringVideoReviewPayload'].map(n=>fn(app,n)).join('\n'),ctx);
 assert.equal(ctx.buildGatheringVideoReviewPayload().prioritySetId,'event');item.review={rating:'like'};ctx.filterGatheringVideoPlaylist();assert.equal(ctx.currentGatheringVideoPlaylist.items.length,0);
 ctx.elements.gatheringVideoView.value='reviewed';ctx.filterGatheringVideoPlaylist();assert.equal(ctx.currentGatheringVideoPlaylist.items.length,1);assert.equal(ctx.currentGatheringVideoPlaylist.items[0].sourceFileId,'one');
 ctx.elements.gatheringEmail.value='other@example.com';assert.throws(()=>ctx.buildGatheringVideoReviewPayload(),/Reload/);
});
test('a saved rating is retained when the follow-up read fails; changing selection cannot overwrite another video',async()=>{
 const original={prioritySetId:'event',sourceFileId:'one'},other={prioritySetId:'other',sourceFileId:'two'};
 const ctx={elements:{gatheringEmail:{value:'me@example.com'}},currentGatheringVideoPlaylist:{allItems:[original]},
 getCurrentGatheringVideoPlaylistItem:()=>original,buildGatheringVideoReviewPayload:()=>({prioritySetId:'event',sourceFileId:'one',email:'me@example.com',rating:'like',notes:'Saved'}),
 postReviewApi:async()=>{ctx.currentGatheringVideoPlaylist={allItems:[other]};ctx.getCurrentGatheringVideoPlaylistItem=()=>other;return {record:{rating:'like',notes:'Saved'}};},
 requestReviewApi:async()=>{throw Error('offline');},normalizeVideoPlaylistItem:x=>x};
 const persist=vm.runInNewContext(fn(app,'persistGatheringVideoReview')+';persistGatheringVideoReview',ctx),result=await persist();
 assert.equal(original.review.rating,'like');assert.equal(other.review,undefined);assert.match(result.queueRefreshError,/Rating saved/);
});

test('an accepted historical handoff stays sent when current file IDs have changed',async()=>{
 const {s,data,gateRecords}=serverHarness();
 data.push(...reviews('old-ready').map(r=>({...r,sourceFileName:'ready'})));
 gateRecords.push({candidateId:'event:old-ready',prioritySetId:'event',sourceFileId:'old-ready',sourceReviewFileId:'old-ready',poetryPleaseHandoff:{status:'sent_to_poetry_please',sourceRecordId:'weaver:video:old-ready',canonicalVideoId:'existing-video'}});
 const d=await s.loadVideoReviewLibrary('me@example.com','event'),item=d.items.find(i=>i.sourceFileId==='ready');
 assert.equal(item.libraryStatus,'sent');assert.equal(item.needsPriorityReview,false);assert.equal(item.sourceFileId,'ready');assert.equal(gateRecords.length,1);
});
