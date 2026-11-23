import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {videoEventCurationMilestone} from '../video_event_milestone.mjs';
const item = (complete = true, excerpts = 1) => ({reviewProgress:{reviewComplete:complete,excerptCount:excerpts}});
test('milestone requires nonempty coverage, complete reviews and excerpts on every included performance',()=>{
 assert.equal(videoEventCurationMilestone([]).curationComplete,false);
 assert.equal(videoEventCurationMilestone([item(),item(false,3)]).curationComplete,false);
 assert.equal(videoEventCurationMilestone([item(),item(true,0)]).curationComplete,false);
 assert.equal(videoEventCurationMilestone([item(),{}]).curationComplete,false);
 const result=videoEventCurationMilestone([item(),item(true,2)]);
 assert.equal(result.curationComplete,true);assert.equal(result.excerptCoveredCount,2);assert.equal(result.missingExcerptCount,0);
});
test('editing, low scores, accepted deliveries and linked format versions neither block the milestone nor get changed',()=>{
 const items=[{...item(),libraryStatus:'sent'},{...item(),libraryStatus:'needs_editing'},
 {...item(),libraryStatus:'review_complete',reviewProgress:{reviewComplete:true,excerptCount:2,candidateScore:6},versionSourceFileIds:['hv','vv']}];
 const before=JSON.stringify(items);assert.equal(videoEventCurationMilestone(items).curationComplete,true);
 assert.equal(videoEventCurationMilestone(items).reviewsCompleteCount,3);assert.equal(JSON.stringify(items),before);
});
const app=fs.readFileSync(new URL('../public/app.js',import.meta.url),'utf8');
const fn=name=>app.match(new RegExp(`^(?:async )?function ${name}\\([\\s\\S]*?^}`, 'm'))[0];
test('completed events have a separate browseable view and remain in library and personal history',()=>{
 const filter=vm.runInNewContext(fn('filterVideoReviewItems')+';filterVideoReviewItems');
 const items=[{poemTitle:'done',eventCurationComplete:true,needsPriorityReview:false,review:{rating:'like'},libraryStatus:'needs_editing'},
 {poemTitle:'pending',eventCurationComplete:false,needsPriorityReview:true}];
 assert.deepEqual(Array.from(filter(items,'complete','').map(i=>i.poemTitle)),['done']);
 assert.deepEqual(Array.from(filter(items,'todo','me@example.com').map(i=>i.poemTitle)),['pending']);
 assert.equal(filter(items,'library','').length,2);assert.equal(filter(items,'reviewed','me@example.com').length,1);
 assert.equal(filter(items,'complete','','','sent').length,0);
});
test('event picker hides complete events from active work without requiring email, but preserves direct links',async()=>{
 const options=[{value:'done'},{value:'pending'}];
 const ctx={elements:{gatheringVideoPrioritySet:{options},gatheringEmail:{value:''},gatheringVideoView:{value:'todo'}},
 requestReviewApi:async()=>({sets:[{id:'done',label:'Done',curationComplete:true,totalCount:2,remainingCount:0},
 {id:'pending',label:'Pending',curationComplete:false,totalCount:1,remainingCount:1}]}),getRequestedVideoSetId:()=>''};
 const refresh=vm.runInNewContext(fn('refreshPriorityVideoSetVisibility')+';refreshPriorityVideoSetVisibility',ctx);
 await refresh();assert.equal(options[0].hidden,true);assert.equal(options[1].hidden,false);
 ctx.elements.gatheringVideoView.value='complete';await refresh();assert.equal(options[0].hidden,false);assert.equal(options[1].hidden,true);assert.match(options[0].textContent,/Curation complete/);
 ctx.elements.gatheringVideoView.value='library';await refresh();assert.ok(options.every(o=>!o.hidden));
 ctx.elements.gatheringVideoView.value='todo';ctx.getRequestedVideoSetId=()=> 'done';await refresh();assert.equal(options[0].hidden,false);
});
