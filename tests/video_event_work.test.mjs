import test from 'node:test';import assert from 'node:assert/strict';import fs from 'node:fs';import vm from 'node:vm';
const server=fs.readFileSync(new URL('../server.mjs',import.meta.url),'utf8'),app=fs.readFileSync(new URL('../public/app.js',import.meta.url),'utf8');
const fn=(s,n)=>s.match(new RegExp(`^(?:async )?function ${n}\\([\\s\\S]*?^}`,'m'))[0];
const policy={minReviewers:5,minScore:8.5};
const summarize=vm.runInNewContext(fn(server,'summarizeVideoEventWork')+';summarizeVideoEventWork',{URL,Map,Set,VIDEO_AUTO_POLICY:policy,videoFamilyForFile:()=>null,extractGoogleDriveFileId:()=>''});
const item=(set,id,n=0,e=0)=>({prioritySetId:set,sourceFileId:id,versionSourceFileIds:[id],needsPriorityReview:n<5||e<1,reviewProgress:{reviewerCount:n,excerptCount:e}});
const pending=(id,n)=>({recordId:'source-'+n,intakeMode:'video',sourceVideoFileId:id});
function report(items,rows=[],progress=[]){const ids=[...new Set(items.map(i=>i.prioritySetId))];return summarize({ok:true,sets:ids.map(id=>({id,label:id,curationComplete:items.filter(i=>i.prioritySetId===id).every(i=>!i.needsPriorityReview)})),items},{ok:true,videos:progress},{ok:true,records:rows},'https://weaver.example','current-time');}

test('largest required task is fresh curation, not extra excerpts on already covered videos',()=>{
 const fresh=Array.from({length:22},(_,i)=>item('new-event','new_'+i));
 const old=Array.from({length:4},(_,i)=>item('covered','old_'+i,5,7));
 const rows=old.flatMap((r,i)=>Array.from({length:6},(_,n)=>pending(r.sourceFileId,i*6+n)));
 const r=report([...fresh,...old],rows,old.map(i=>({sourceVideoFileId:i.sourceFileId,approvedCount:1})));
 assert.equal(r.largestTask.eventId,'new-event');assert.equal(r.largestTask.videoCount,22);
 assert.equal(r.events[0].reviewSlotsNeeded,110);assert.equal(r.events[0].videosMissingExcerpts,22);
 const covered=r.events.find(e=>e.eventId==='covered');assert.equal(covered.pendingExcerptCount,24);assert.equal(covered.additionalPendingExcerptCount,24);assert.equal(covered.videosNeedingExcerptApproval,0);
 assert.equal(r.tasks.length,1);assert.equal(new URL(r.largestTask.workUrls[0]).searchParams.get('videoSet'),'new-event');assert.ok(!JSON.stringify(r).includes('file://'));
});

test('exact linked versions count once and ambiguous parents are not silently assigned',()=>{
 const one={...item('one','representative',5,3),versionSourceFileIds:['raw','finished']};
 const r=report([one],[pending('raw',1),pending('raw',1),pending('finished',2)],[{sourceVideoFileId:'raw',approvedCount:2}]);
 assert.equal(r.events[0].videoCount,1);assert.equal(r.events[0].pendingExcerptCount,2);assert.equal(r.events[0].videosWithApprovedExcerpts,1);assert.equal(r.tasks.length,0);
 const conflict=report([one,{...item('other','separate',5,3),versionSourceFileIds:['raw']}],[pending('raw',1)]);
 assert.equal(conflict.unassignedPendingExcerptCount,1);assert.ok(conflict.events.every(e=>e.pendingExcerptCount===0));
});

test('approval assignments contain only videos needing approval and respect existing 50-file task limit',()=>{
 const ids=Array.from({length:51},(_,i)=>'source_file_'+String(i).padStart(5,'0'));
 const r=report([{...item('event','group',5,51),versionSourceFileIds:ids}],ids.map((id,n)=>pending(id,n)));
 assert.equal(r.events[0].videosNeedingExcerptApproval,1);assert.equal(r.events[0].requiredApprovalCandidateCount,51);assert.equal(r.events[0].excerptApprovalUrls.length,2);
 assert.equal(r.events[0].excerptApprovalUrls.reduce((n,u)=>n+new URL(u).searchParams.get('videoFiles').split(',').length,0),51);
 assert.ok(r.events[0].excerptApprovalUrls.every(u=>new URL(u).searchParams.get('task')==='approve-video-excerpts'));
});

test('unavailable sources remain errors, never zero-work reports',()=>{
 assert.throws(()=>summarize({ok:false},{ok:true,videos:[]},{ok:true,records:[]},'https://weaver.example','now'),/not verified/);
 const empty=summarize({ok:true,sets:[],items:[]},{ok:true,videos:[]},{ok:true,records:[]},'https://weaver.example','now');assert.equal(empty.largestTask,null);assert.equal(empty.events.length,0);
});

test('exact new event link opens without global availability or local options; no email opens a readable library',async()=>{
 for(const email of ['', 'person@example.com']){
  const options=[],calls=[],select={options,add:option=>options.push(option)};
  const s={contributorShellParams:new URLSearchParams({videoSet:'new-event'}),elements:{gatheringVideoPrioritySet:select,gatheringVideoPlaylistUrl:{value:'old-url'},gatheringVideoView:{value:'reviewed'},gatheringEmail:{value:email}},Option:function(label,value){return{textContent:label,value};},setActiveModule:v=>calls.push(v),setGatheringMode:(v,options)=>{assert.equal(options.refreshVideoSets,false);calls.push(v);},refreshPriorityVideoSetVisibility:()=>{throw Error('Global request must not run');},loadGatheringVideoPlaylist:async()=>{assert.equal(select.value,'new-event');calls.push('exact-load');},setVideoReviewSaveStatus:e=>{throw Error(e);}};
  await vm.runInNewContext(fn(app,'openRequestedVideoReviewTask')+';openRequestedVideoReviewTask',s)();
  assert.equal(s.elements.gatheringVideoView.value,email?'todo':'library');assert.equal(s.elements.gatheringVideoPlaylistUrl.value,'');assert.equal(calls.at(-1),'exact-load');
 }
});
