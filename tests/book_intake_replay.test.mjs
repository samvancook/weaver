import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const source=fs.readFileSync(new URL('../server.mjs',import.meta.url),'utf8');
function fn(name) {
  const match=source.match(new RegExp(`^async function ${name}\\([\\s\\S]*?^}\\n`,'m'));
  assert.ok(match,name);return match[0];
}
test('an already-saved book submission requests an intake-only confirmation',async()=>{
  const calls=[];
  const c={String,Error,excerptStorageMode:'firestore',appVersion:'test',
    cleanSheetWhitespace:v=>String(v||'').trim(),findExistingBookIntakeRow:async()=>12,
    buildExcerptRuntimeRecord:p=>p,
    syncWeaverRuntimeDb:async(action,payload)=>{calls.push({action,payload});return {ok:true,count:1};}
  };
  vm.createContext(c);vm.runInContext(fn('shadowExcerptRecordsToRuntime')+'\n'+fn('appendExcerptGatheringRow'),c);
  const result=await c.appendExcerptGatheringRow({mode:'book',recordId:'source-one'});
  assert.equal(result.alreadySaved,true);assert.equal(result.rowNumber,12);
  assert.equal(calls.length,1);assert.equal(calls[0].payload.intakeOnly,true);
  await c.shadowExcerptRecordsToRuntime([{recordId:'source-one',reviewDecision:'ACCEPT'}]);
  assert.equal(calls[1].payload.intakeOnly,false,'Normal review updates must remain enabled');
});
