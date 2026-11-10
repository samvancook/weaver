import test from'node:test';import assert from'node:assert/strict';import fs from'node:fs';import vm from'node:vm';
const source=fs.readFileSync(new URL('../public/app.js',import.meta.url),'utf8');const fn=name=>source.match(new RegExp(`^(?:async )?function ${name}\\([\\s\\S]*?^}`, 'm'))[0];
test('workspace Review videos loads the selected durable event and clears the legacy form without a catalog dependency',async()=>{
 const calls=[],s={elements:{videoCurationEvent:{value:'Finals'},gatheringVideoPlaylistUrl:{value:'legacy-form'},gatheringVideoPrioritySet:{value:'',options:[{value:'mpmu'}]}},
 videoCurationCandidates:[{sourceEvent:'Finals',prioritySetId:'mpmu'}],currentGatheringVideoPlaylist:null,setActiveModule:x=>calls.push(x),setGatheringMode:x=>calls.push(x),setVideoReviewSaveStatus:x=>calls.push(x),refreshPriorityVideoSetVisibility:async()=>{},loadGatheringVideoPlaylist:async()=>{calls.push('loaded');s.currentGatheringVideoPlaylist={prioritySetId:'mpmu'};}};
 const open=vm.runInNewContext(fn('openVideoReviewFromWorkspace')+';openVideoReviewFromWorkspace',s);await open();assert.equal(s.elements.gatheringVideoPlaylistUrl.value,'');assert.equal(s.elements.gatheringVideoPrioritySet.value,'mpmu');assert.ok(calls.includes('loaded'));assert.match(calls.at(-1),/without excerpts/);
});
test('Save Rating & Excerpts saves a loaded video without quotes, rather than routing it as legacy intake',async()=>{
 const calls=[],s={getSelectedGatheringMode:()=> 'video',videoReviewSaving:false,setVideoReviewSaving:x=>calls.push(['busy',x]),buildGatheringPayload:()=>({mode:'video',prioritySetId:'mpmu',sourceFileId:'tilt',quotes:[]}),setVideoReviewSaveStatus:x=>calls.push(x),persistGatheringVideoReview:async()=>calls.push('rating'),resetGatheringAfterSubmit:()=>calls.push('advance'),setGatheringEmailWarning:()=>{}};
 const submit=vm.runInNewContext(fn('submitGathering')+';submitGathering',s);await submit();assert.equal(calls.filter(x=>x==='rating').length,1);assert.ok(calls.includes('advance'));assert.match(calls.find(x=>typeof x==='string'&&x.startsWith('Video rating saved')),/No excerpts/);assert.equal(calls.at(-1)[1],false);
 s.buildGatheringPayload=()=>({mode:'video',prioritySetId:'',quotes:[]});calls.length=0;await submit();assert.equal(calls.filter(x=>x==='rating').length,0);assert.match(calls.find(x=>typeof x==='string'&&x.includes('not saved')),/Choose the video set/);assert.ok(!calls.includes('advance'));
});
test('timeout readback confirms only the exact newly saved reviewer/file/rating and never repeats a write',async()=>{
 const item={sourceFileId:'tilt'},s={Date,currentGatheringVideoPlaylist:{},getCurrentGatheringVideoPlaylistItem:()=>item,buildGatheringVideoReviewPayload:()=>({prioritySetId:'mpmu',sourceFileId:'tilt',email:'sam@buttonpoetry.com',rating:'like',notes:'My note'}),postReviewApi:async()=>{throw Object.assign(new Error('timeout'),{saveOutcomeUnknown:true});},requestReviewApi:async()=>({items:[{sourceFileId:'tilt',review:{rating:'like',notes:'My note',updatedAt:new Date().toISOString()}}]})};
 const persist=vm.runInNewContext(fn('persistGatheringVideoReview')+';persistGatheringVideoReview',s);const result=await persist();assert.equal(result.confirmedByReadback,true);assert.equal(item.review.rating,'like');
 s.requestReviewApi=async()=>({items:[{sourceFileId:'other',review:{rating:'like',notes:'My note',updatedAt:new Date().toISOString()}}]});await assert.rejects(persist(),/timeout/);
 s.requestReviewApi=async()=>({items:[{sourceFileId:'tilt',review:{rating:'like',notes:'My note',updatedAt:'2025-01-01T00:00:00Z'}}]});await assert.rejects(persist(),/timeout/);
});
test('rating failure leaves entered fields untouched and clears the pending state',async()=>{
 const calls=[],s={videoReviewSaving:false,currentGatheringVideoPlaylist:{},setVideoReviewSaving:v=>calls.push(v),setVideoReviewSaveStatus:m=>calls.push(m),persistGatheringVideoReview:async()=>{throw new Error('Unavailable');},resetGatheringAfterSubmit:()=>{throw Error('Do not clear the draft');}};
 const save=vm.runInNewContext(fn('saveGatheringVideoRatingAndNext')+';saveGatheringVideoRatingAndNext',s);await save();assert.equal(calls.at(-1),false);assert.match(calls.find(x=>typeof x==='string'&&x.includes('not saved')),/Unavailable/);
});
test('video request timeout is bounded and reported as uncertain, with no automatic replay',async()=>{
 let aborted,posts=0;const s={AbortController,setTimeout:cb=>{aborted=cb;return 1;},clearTimeout:()=>{},fetch:(_url,options)=>{posts++;return new Promise((_,reject)=>options.signal.addEventListener('abort',()=>reject(Object.assign(new Error('aborted'),{name:'AbortError'}))));}};
 const post=vm.runInNewContext(fn('postReviewApi')+';postReviewApi',s),promise=post('/api/intake/video-review',{});aborted();await assert.rejects(promise,e=>e.saveOutcomeUnknown&&/no automatic retry/.test(e.message));assert.equal(posts,1);
});
test('an excerpt failure after a confirmed rating does not falsely report that the rating was lost',async()=>{
 const messages=[],s={getSelectedGatheringMode:()=> 'video',videoReviewSaving:false,setVideoReviewSaving:()=>{},buildGatheringPayload:()=>({mode:'video',prioritySetId:'mpmu',quotes:['one']}),setVideoReviewSaveStatus:m=>messages.push(m),setStatus:()=>{},persistGatheringVideoReview:async()=>({ok:true}),postReviewApi:async()=>{throw Error('Excerpt unavailable');}};
 const submit=vm.runInNewContext(fn('submitGathering')+';submitGathering',s);await submit();assert.match(messages.at(-1),/^Rating saved\./);assert.match(messages.at(-1),/Excerpt unavailable/);
});
test('a direct main-site task link opens the selected Weaver set without loading an old form',async()=>{
 const calls=[],s={contributorShellParams:new URLSearchParams({videoSet:'mpmu-2026-finals',video:'tilt'}),elements:{gatheringVideoPrioritySet:{options:[{value:'mpmu-2026-finals'}]},gatheringVideoPlaylistUrl:{value:'old-form'}},setActiveModule:()=>{},setGatheringMode:()=>{},refreshPriorityVideoSetVisibility:async()=>{},loadGatheringVideoPlaylist:async()=>calls.push('loaded'),setVideoReviewSaveStatus:m=>calls.push(m)};
 const open=vm.runInNewContext(fn('openRequestedVideoReviewTask')+';openRequestedVideoReviewTask',s);await open();assert.equal(s.elements.gatheringVideoPrioritySet.value,'mpmu-2026-finals');assert.equal(s.elements.gatheringVideoPlaylistUrl.value,'');assert.deepEqual(calls,['loaded']);
});
test('a direct video task never falls through to another video when its own review is saved or missing',async()=>{
 for(const mode of ['saved','missing','pending']){
  const displayed=[],messages=[],items=[{sourceFileId:'other',author:'Other',poemTitle:'Other'}];if(mode!=='missing')items.push({sourceFileId:'tilt',author:'Keaghan',poemTitle:'Tilt',review:mode==='saved'?{rating:'like'}:null});
  const s={videoPlaylistGeneration:0,resolveVideoPlaylistRequest:(prioritySetId,formUrl)=>({prioritySetId,formUrl}),filterGatheringVideoPlaylist:()=>{},renderVideoLibraryResults:()=>{},currentGatheringVideoPlaylist:null,elements:{gatheringVideoPlaylistUrl:{value:''},gatheringVideoPrioritySet:{value:'mpmu'},gatheringEmail:{value:'sam@buttonpoetry.com'}},contributorShellParams:new URLSearchParams({video:'tilt'}),setStatus:m=>messages.push(m),updateGatheringVideoPlaylistUi:()=>{},requestReviewApi:async()=>({prioritySetId:'mpmu',items,eventName:'Finals'}),normalizeVideoPlaylistItem:x=>x,cleanSheetWhitespace:x=>String(x||'').trim(),clearGatheringVideoPlaylistItemFields:()=>{},applyGatheringVideoPlaylistItem:i=>displayed.push(i.sourceFileId),getSelectedGatheringMode:()=> 'video',setGatheringMode:()=>{},refreshPriorityVideoSetVisibility:async()=>{}};
  s.elements.gatheringVideoView={value:'todo'};s.videoReviewView=()=> 'todo';const load=vm.runInNewContext(['filterVideoReviewItems','filterGatheringVideoPlaylist','loadGatheringVideoPlaylist'].map(fn).join('\n')+';loadGatheringVideoPlaylist',s);await load();if(mode==='pending')assert.deepEqual(displayed,['tilt']);else {assert.equal(displayed.length,0);assert.match(messages.at(-1),mode==='saved'?/already saved/:/No other video/);}
 }
});
test('the legacy manual excerpt button remains available while video saves are idle',()=>{
 const s={currentGatheringVideoPlaylist:null,videoReviewSaving:false,elements:{submitGathering:{disabled:true}},getSelectedGatheringMode:()=> 'video',setGatheringVideoPlaylistStatus:()=>{},cleanSheetWhitespace:x=>String(x||'').trim()};
 const update=vm.runInNewContext(fn('updateGatheringVideoPlaylistUi')+';updateGatheringVideoPlaylistUi',s);update();assert.equal(s.elements.submitGathering.disabled,false);s.videoReviewSaving=true;update();assert.equal(s.elements.submitGathering.disabled,true);
});
