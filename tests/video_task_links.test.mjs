import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import vm from 'node:vm';
const app=readFileSync(new URL('../public/app.js',import.meta.url),'utf8');
const fn=name=>app.match(new RegExp(`^(?:async )?function ${name}\\([\\s\\S]*?^}`, 'm'))[0];
test('event labels explicitly distinguish personal queue, personal history and total library counts',async()=>{
 for(const [view,label] of [['todo','3 awaiting your review'],['reviewed','2 reviewed by you'],['library','16 videos total']]){
  const option={value:'mpmu'},s={elements:{gatheringVideoPrioritySet:{options:[option]},gatheringEmail:{value:'sam@buttonpoetry.com'},gatheringVideoView:{value:view}},getRequestedVideoSetId:()=>'',requestReviewApi:async()=>({sets:[{id:'mpmu',label:'MPMU',remainingCount:3,reviewedCount:2,totalCount:16}]})};
  await vm.runInNewContext(fn('refreshPriorityVideoSetVisibility')+';refreshPriorityVideoSetVisibility',s)();assert.equal(option.textContent,`MPMU — ${label}`);
 }
});
test('without reviewer email the event labels request identity instead of suggesting a personal count',async()=>{
 const option={value:'mpmu'},s={elements:{gatheringVideoPrioritySet:{options:[option]},gatheringEmail:{value:''},gatheringVideoView:{value:'todo'}},getRequestedVideoSetId:()=>'',requestReviewApi:async()=>({sets:[{id:'mpmu',label:'MPMU',remainingCount:0}]})};
 await vm.runInNewContext(fn('refreshPriorityVideoSetVisibility')+';refreshPriorityVideoSetVisibility',s)();assert.match(option.textContent,/enter your email/);assert.equal(option.hidden,false);
});
test('same-site task URL retains its exact event while Google Form URLs remain usable',()=>{
 const origin='https://weaver.buttonpoetry.com',s={URL,window:{location:{origin}}},resolve=vm.runInNewContext(fn('resolveVideoPlaylistRequest')+';resolveVideoPlaylistRequest',s);
 assert.equal(resolve('',origin+'/?videoSet=mpmu-2026-finals').prioritySetId,'mpmu-2026-finals');assert.equal(resolve('','https://forms.gle/legacy').formUrl,'https://forms.gle/legacy');assert.throws(()=>resolve('','https://other.example/?videoSet=mpmu-2026-finals'),/own Weaver site/);
});
test('a selected exact task does not constrain a different event selected afterward',()=>{
 const items=[{sourceFileId:'different',author:'Poet',poemTitle:'Poem',needsPriorityReview:true}],s={contributorShellParams:new URLSearchParams({videoSet:'first',video:'tilt'}),currentGatheringVideoPlaylist:{prioritySetId:'other',allItems:items,items,index:0,reviewerEmail:'sam'},elements:{},videoReviewView:()=> 'todo',applyGatheringVideoPlaylistItem:()=>{},clearGatheringVideoPlaylistItemFields:()=>{},renderVideoLibraryResults:()=>{},updateGatheringVideoPlaylistUi:()=>{}};
 vm.runInNewContext(fn('filterVideoReviewItems')+'\n'+fn('filterGatheringVideoPlaylist')+';filterGatheringVideoPlaylist()',s);assert.equal(s.currentGatheringVideoPlaylist.items.length,1);
});

const origin='https://footage-bridge---weaver-iasomqy6ja-uc.a.run.app';
test('Helper exposes a task URL only for a currently accepted saved connection',()=>{
 const helper=readFileSync(new URL('../integrations/footage-helper/WeaverBridge.gs',import.meta.url),'utf8');
 const project=vm.runInNewContext(helper+';weaverReceiptForUi_', {WEAVER_REVIEWED_TEST_URL:origin,encodeURIComponent});
 const record={eventId:'inventory-event',prioritySetId:'mpmu-2026-finals',manifest:{closed:true},transferComplete:true,status:'linked_existing_event'};
 assert.equal(project(record).curationTaskUrl,origin+'/?videoSet=mpmu-2026-finals');
 for(const extra of [{preview:true},{transferComplete:false},{status:'inventory_changed'}]) assert.equal(project({...record,...extra}).curationTaskUrl,'');
});
