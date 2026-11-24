import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const app=fs.readFileSync(new URL('../public/app.js',import.meta.url),'utf8');
const body=app.match(/^async function submitExcerptSet\([\s\S]*?^}\n/m)[0];
function harness() {
  const calls=[];
  const c={isSaving:false,Number,Error,Array,normalizeDecision:x=>x,filterChangedUpdates:x=>x,
    setStatus:s=>calls.push(['status',s]),setSubmitState:v=>{c.isSaving=v;},
    requestBatchSave:async()=>{calls.push(['save']);return {ok:true,savedCount:1,excerptHandoffs:{ok:true,queuedVideoCount:1}};}
  };
  vm.createContext(c);vm.runInContext(body,c);
  const draft={sourceRow:2,recordId:'one',reviewDecision:'accept_skip_graphic'};
  const options={sourceExcerpts:[draft],collectUpdates:()=>[draft],idleLabel:'Submit Decisions',progressLabel:'Saving',
    afterSaveOptimistic:()=>calls.push(['optimistic']),
    reloadRequest:async()=>{throw Error('read timed out');},afterReload:async()=>{}};
  return {c,calls,options,draft,status:()=>calls.filter(x=>x[0]==='status').at(-1)[1]};
}
test('confirmed excerpt decisions survive refresh failure without being labelled a failed save',async()=>{
  const h=harness();await h.c.submitExcerptSet(h.options);
  assert.match(h.status(),/^Saved 1 changed decisions.*delivery queued.*Queue refresh failed.*Do not resubmit/);
  assert.equal(h.calls.filter(x=>x[0]==='save').length,1);assert.equal(h.c.isSaving,false);
});
test('an uncertain save retains the draft and makes no optimistic removal or automatic retry',async()=>{
  const h=harness();h.c.requestBatchSave=async()=>{h.calls.push(['save']);throw Error('connection lost');};
  await h.c.submitExcerptSet(h.options);
  assert.match(h.status(),/^Submit not confirmed:.*Some decisions may already be saved.*No automatic retry/);
  assert.equal(h.calls.filter(x=>x[0]==='save').length,1);
  assert.ok(!h.calls.some(x=>x[0]==='optimistic'));assert.equal(h.draft.reviewDecision,'accept_skip_graphic');
  assert.equal(h.c.isSaving,false);
});
