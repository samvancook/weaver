import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const app = fs.readFileSync(new URL('../public/app.js', import.meta.url), 'utf8');
const server = fs.readFileSync(new URL('../server.mjs', import.meta.url), 'utf8');
function fn(source, name) {
  const match = source.match(new RegExp(`^(?:async )?function ${name}\\([\\s\\S]*?^}\\n`, 'm'));
  assert.ok(match, name);
  return match[0];
}
const update = (id='one') => ({recordId:id,storageTarget:'firestore',pigCompletionId:id,graphicsRequestId:`request-${id}`,qcDecision:'approve'});
function ui(result={ok:true,savedCount:1,runtimeDb:{ok:true,written:1}}) {
  const calls=[];
  const c={Set,Number,Error,isSaving:false,currentGraphicsRecords:[update()],currentGraphicsAssetMatches:new Map(),
    elements:{graphicsBookSelect:{value:'Book'}}, getSelectedGraphicsMode:()=> 'qc',
    collectGraphicsQcUpdates:()=>[update()],filterChangedGraphicsQcUpdates:x=>x,
    setStatus:s=>calls.push(['status',s]), setSubmitState:busy=>{c.isSaving=busy;},
    adminFetch:async()=>{calls.push(['save']);return {ok:true,json:async()=>result};},
    isGraphicsQcSweepSelection:()=>false,renderGraphicsRecords:r=>calls.push(['render',r.length]),
    applyOptimisticGraphicsBookCountUpdate:()=>{}, loadGraphicsBooks:async()=>{},loadGraphicsRecords:async()=>{}
  };
  vm.createContext(c);vm.runInContext(fn(app,'submitGraphicsQc'),c);
  return {c,calls,status:()=>calls.filter(x=>x[0]==='status').at(-1)[1]};
}

test('QC rejects an incomplete acknowledgment without removing the unsaved cards or retrying',async()=>{
  for(const result of [{ok:true,savedCount:0,runtimeDb:{ok:true,written:1}},{ok:true,savedCount:1,runtimeDb:{ok:true,written:0}},{ok:true,savedCount:1}]) {
    const h=ui(result);await h.c.submitGraphicsQc();
    assert.equal(h.c.currentGraphicsRecords.length,1);
    assert.equal(h.calls.filter(x=>x[0]==='save').length,1);
    assert.match(h.status(),/not confirmed.*Some may have saved/);
    assert.equal(h.c.isSaving,false);
  }
});
test('a refresh error after confirmed QC saves reports saved decisions separately',async()=>{
  const h=ui();h.c.loadGraphicsBooks=async()=>{throw Error('temporary read failure');};
  await h.c.submitGraphicsQc();
  assert.match(h.status(),/^Saved 1 QC decisions.*Refresh failed.*Do not resubmit/);
  assert.equal(h.c.currentGraphicsRecords.length,0);
  assert.equal(h.calls.filter(x=>x[0]==='save').length,1);
});
test('sheet synchronization failure is visible without denying a confirmed ledger save',async()=>{
  const h=ui({ok:true,savedCount:1,runtimeDb:{ok:true,written:1},sheetSync:{ok:false,error:'Sheet unavailable'}});
  await h.c.submitGraphicsQc();assert.match(h.status(),/^Saved 1 QC decisions.*Sheet synchronization needs attention.*Sheet unavailable/);
});
test('a second QC submit while the first is pending sends nothing',async()=>{
  const h=ui();let finish;
  h.c.adminFetch=()=>{h.calls.push(['save']);return new Promise(resolve=>{finish=resolve;});};
  const saving=h.c.submitGraphicsQc();await h.c.submitGraphicsQc();
  assert.equal(h.calls.filter(x=>x[0]==='save').length,1);
  finish({ok:true,json:async()=>({ok:true,savedCount:1,runtimeDb:{ok:true,written:1}})});
  await saving;assert.equal(h.c.isSaving,false);
});
test('QC preflight rejects malformed and duplicate targets before any service call',async()=>{
  const calls=[];
  const c={Set,Number,String,Error,cleanSheetWhitespace:v=>String(v||'').trim(),
    ensurePigCompletedGraphicsSheetServer:async()=>calls.push('ensure'),
    normalizeGraphicsQcDecision:v=>['approve','reject','replace'].includes(v)?'APPROVE':'',};
  vm.createContext(c);vm.runInContext(fn(server,'validateGraphicsQcBatch')+'\n'+fn(server,'saveGraphicsQcToSheets'),c);
  for(const batch of [[update(),update()], [update(),{...update('two'),graphicsRequestId:''}],
    [update(),{...update('two'),storageTarget:'unknown'}], [update(),{...update('two'),storageTarget:'pig_sheet',sheetRow:'2junk'}],
    [update(),{...update('two'),qcDecision:''}]]) {
    await assert.rejects(c.saveGraphicsQcToSheets(batch));assert.deepEqual(calls,[]);
  }
  const normalized=c.validateGraphicsQcBatch([{...update(),storageTarget:'sheet_cleanup',sheetRow:2}]);
  assert.equal(normalized[0].storageTarget,'cleanup_sheet');
});
