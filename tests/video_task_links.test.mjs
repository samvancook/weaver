import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {runInNewContext} from 'node:vm';
const source=readFileSync(new URL('../public/app.js',import.meta.url),'utf8');
const fn=name=>source.match(new RegExp(`^(?:async )?function ${name}\\([\\s\\S]*?^}`, 'm'))[0];
const origin='https://footage-bridge---weaver-iasomqy6ja-uc.a.run.app';
function harness({id='mpmu-2026-finals',available=true,reviewed=false,returnedId=id}={}){
 const calls=[],status=[];
 const select={value:'',options:[]};
 const elements={gatheringVideoPrioritySet:select,gatheringVideoPlaylistUrl:{value:'old-form'},gatheringEmail:{value:'reviewer@example.com'}};
 const ctx={URL,Array,contributorShellParams:new URLSearchParams({videoSet:id}),cleanSheetWhitespace:x=>String(x||'').trim(),window:{location:{origin}},elements,
  currentGatheringVideoPlaylist:null,setActiveModule:name=>calls.push(['module',name]),setGatheringMode:name=>calls.push(['mode',name]),
  refreshPriorityVideoSetVisibility:async()=>{select.options=available?[{value:id}]:[];},
  requestReviewApi:async(path,params)=>{calls.push(['read',path,params]);return {ok:true,prioritySetId:returnedId,eventName:'MPMU Finals',items:[{sourceFileId:'video',author:'Poet',poemTitle:'Poem',review:reviewed?{rating:'like'}:null}]};},
  normalizeVideoPlaylistItem:x=>x,clearGatheringVideoPlaylistItemFields:()=>calls.push(['clear']),applyGatheringVideoPlaylistItem:x=>calls.push(['display',x.sourceFileId]),
  getSelectedGatheringMode:()=> 'video',updateGatheringVideoPlaylistUi:()=>{},setStatus:s=>status.push(s)};
 const api=runInNewContext([fn('getRequestedVideoSetId'),fn('resolveVideoPlaylistRequest'),fn('loadGatheringVideoPlaylist'),fn('openRequestedVideoTask')].join('\n')+';({openRequestedVideoTask,resolveVideoPlaylistRequest,getRequestedVideoSetId})',ctx);
 return {ctx,api,calls,status};
}
test('event link opens the exact curation module and playlist with the current reviewer, using read requests only',async()=>{
 const {ctx,api,calls}=harness();await api.openRequestedVideoTask();
 assert.equal(ctx.elements.gatheringVideoPrioritySet.value,'mpmu-2026-finals');
 assert.equal(ctx.elements.gatheringVideoPlaylistUrl.value,'');
 assert.equal(ctx.currentGatheringVideoPlaylist.prioritySetId,'mpmu-2026-finals');
 assert.deepEqual(calls.filter(x=>x[0]==='module'),[['module','gathering']]);
 assert.deepEqual(calls.filter(x=>x[0]==='mode'),[['mode','video']]);
 const read=calls.find(x=>x[0]==='read');assert.equal(read[2].reviewerEmail,'reviewer@example.com');assert.equal(read[2].prioritySetId,'mpmu-2026-finals');
});
test('a completed task opens its completion state without restarting reviews',async()=>{
 const {ctx,api,calls}=harness({reviewed:true});await api.openRequestedVideoTask();
 assert.equal(ctx.currentGatheringVideoPlaylist.completed,true);assert.equal(ctx.currentGatheringVideoPlaylist.items.length,0);
 assert.ok(!calls.some(x=>x[0]==='display'));assert.equal(ctx.currentGatheringVideoPlaylist.reviewedCount,1);
});
test('unknown, malformed and mismatched task links show errors rather than opening another event',async()=>{
 for(const config of [{available:false},{id:'../../bad'},{id:''},{returnedId:'other-event'}]){
  const {ctx,api,status}=harness(config);await api.openRequestedVideoTask();
  assert.equal(ctx.currentGatheringVideoPlaylist,null);assert.match(status.at(-1),/could not be opened|load failed/);
 }
});
test('a task link keeps a completed event selectable while other completed queues stay hidden',async()=>{
 const option={value:'mpmu-2026-finals'},other={value:'other-event'};
 const select={value:option.value,options:[option,other],querySelector:()=>({append:()=>{}})};
 const ctx={Array,contributorShellParams:new URLSearchParams({videoSet:option.value}),cleanSheetWhitespace:x=>String(x||'').trim(),
  elements:{gatheringVideoPrioritySet:select,gatheringEmail:{value:'reviewer@example.com'}},
  requestReviewApi:async()=>({sets:[{id:option.value,completed:true},{id:other.value,completed:true}]})};
 const refresh=runInNewContext(fn('getRequestedVideoSetId')+'\n'+fn('refreshPriorityVideoSetVisibility')+';refreshPriorityVideoSetVisibility',ctx);
 await refresh();assert.equal(option.hidden,false);assert.equal(option.disabled,false);assert.equal(select.value,option.value);assert.equal(other.hidden,true);
});
test('the playlist field understands a same-site Weaver task URL and retains Google Form support',()=>{
 const {api}=harness();
 const linked=api.resolveVideoPlaylistRequest('',origin+'/?videoSet=mpmu-2026-finals');assert.equal(linked.prioritySetId,'mpmu-2026-finals');assert.equal(linked.formUrl,'');
 const form='https://forms.gle/existing';assert.equal(api.resolveVideoPlaylistRequest('',form).formUrl,form);
 assert.throws(()=>api.resolveVideoPlaylistRequest('','https://other.example/?videoSet=mpmu-2026-finals'),/own Weaver site/);
});
test('Helper exposes a task URL only for a currently accepted saved connection',()=>{
 const helper=readFileSync(new URL('../integrations/footage-helper/WeaverBridge.gs',import.meta.url),'utf8');
 const project=runInNewContext(helper+';weaverReceiptForUi_', {WEAVER_REVIEWED_TEST_URL:origin,encodeURIComponent});
 const record={eventId:'inventory-event',prioritySetId:'mpmu-2026-finals',manifest:{closed:true},transferComplete:true,status:'linked_existing_event'};
 assert.equal(project(record).curationTaskUrl,origin+'/?videoSet=mpmu-2026-finals');
 for(const extra of [{preview:true},{transferComplete:false},{status:'inventory_changed'}]) assert.equal(project({...record,...extra}).curationTaskUrl,'');
});
