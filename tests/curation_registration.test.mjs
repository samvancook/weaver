import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const scope={};vm.createContext(scope);vm.runInContext(fs.readFileSync(new URL('../integrations/footage-helper/CurationRegistration.gs',import.meta.url),'utf8'),scope);
function fixture(){
 const audit=[],rows=[];let serial=0;
 const context={eventName:'PAGES',taskUrl:'https://weaver.test/?videoSet=pages',sourceEventId:'receipt-1',prioritySetId:'pages',sourceSnapshot:{existingCurationLinks:[]}};
 const state={audit,rows,context,projection:true,failAudit:'',failAppend:false,failWithdraw:false};
 const store={context:()=>structuredClone(context),history:()=>structuredClone(audit),active:()=>rows.filter(r=>!r.withdrawn).map(r=>({operationId:r.op.operationId,url:r.op.taskUrl})),id:()=> 'sandbox',url:()=> 'https://sheets.test',actor:()=> 'sam@test',now:()=> '2026-11-04T10:00:00Z',uuid:()=> 'op-'+(++serial),fingerprint:JSON.stringify,
 audit:r=>{if(r.state===state.failAudit)throw Error('audit unavailable');audit.push(structuredClone(r));},
 append:op=>{rows.push({op:structuredClone(op),withdrawn:false});if(state.failAppend)throw Error('lost response');return {row:rows.length+1};},
 verify:(op,withdrawn)=>state.projection&&rows.some(r=>r.op.operationId===op.operationId&&!r.changed&&r.withdrawn===withdrawn),
 withdraw:op=>{rows.find(r=>r.op.operationId===op.operationId).withdrawn=true;if(state.failWithdraw)throw Error('lost withdrawal response');return {row:2};}};
 const run=(action,extra={})=>scope.curationRegistrationEngine_(store,action,{eventName:'PAGES',...extra});
 const register=()=>run('register',{previewKey:run('preview').previewKey});
 return {...state,state,store,run,register};
}
test('intent precedes source, verified registration is duplicate safe; withdrawal keeps full history',()=>{
 const f=fixture();const r=f.register();assert.equal(r.state,'VERIFIED');assert.deepEqual(f.audit.map(x=>x.state),['INTENT','SOURCE_WRITTEN','VERIFIED']);
 assert.equal(f.register().state,'EXISTING_LINK_PRESERVED');assert.equal(f.rows.length,1);
 assert.equal(f.run('withdraw',{operationId:r.operationId}).state,'WITHDRAWN');assert.equal(f.rows[0].withdrawn,true);assert.equal(f.audit.length,6);
 assert.equal(f.run('withdraw',{operationId:r.operationId}).state,'WITHDRAWN');assert.equal(f.audit.length,6);
});
test('failed intent cannot write source',()=>{const f=fixture();f.state.failAudit='INTENT';assert.throws(()=>f.register(),/audit unavailable/);assert.equal(f.rows.length,0);});
test('stale preview refuses and current production links are protected without sandbox baseline',()=>{
 const f=fixture(),p=f.run('preview');f.context.taskUrl+='&changed=1';assert.throws(()=>f.run('register',{previewKey:p.previewKey}),/preview changed/);assert.equal(f.rows.length,0);
 f.context.sourceSnapshot.existingCurationLinks.push({url:'https://forms.test/legacy'});assert.equal(f.run('preview').canRegister,false);assert.equal(f.register().state,'EXISTING_LINK_PRESERVED');assert.equal(f.rows.length,0);
});
for(const failure of ['failAppend','failAudit'])test('ambiguous '+failure+' never appends twice and read-only reconciliation verifies outcome',()=>{
 const f=fixture();if(failure==='failAudit')f.state.failAudit='SOURCE_WRITTEN';else f.state.failAppend=true;
 const r=f.register();assert.equal(r.state,'RECONCILIATION_REQUIRED');assert.equal(f.register().state,'RECONCILIATION_REQUIRED');assert.equal(f.rows.length,1);
 f.state.failAudit='';assert.equal(f.run('reconcile',{operationId:r.operationId}).state,'VERIFIED');assert.equal(f.rows.length,1);
});
test('projection failure remains unresolved until source AND view verify',()=>{const f=fixture();f.state.projection=false;const r=f.register();assert.equal(r.state,'RECONCILIATION_REQUIRED');assert.equal(f.run('reconcile',{operationId:r.operationId}).state,'RECONCILIATION_REQUIRED');f.state.projection=true;assert.equal(f.run('reconcile',{operationId:r.operationId}).state,'VERIFIED');});
test('changed owned row or another registration is preserved on withdrawal',()=>{
 for(const change of ['edit','another']){const f=fixture(),r=f.register();if(change==='edit')f.rows[0].changed=true;else f.rows.push({op:{operationId:'manual',taskUrl:'other'},withdrawn:false});assert.equal(f.run('withdraw',{operationId:r.operationId}).state,'CHANGED_RECORD_PRESERVED');assert.equal(f.rows[0].withdrawn,false);}
});
test('uncertain withdrawal reconciles without another mutation',()=>{const f=fixture(),r=f.register();f.state.failWithdraw=true;assert.equal(f.run('withdraw',{operationId:r.operationId}).state,'RECONCILIATION_REQUIRED');assert.equal(f.run('withdraw',{operationId:r.operationId}).state,'RECONCILIATION_REQUIRED');assert.equal(f.run('reconcile',{operationId:r.operationId}).state,'WITHDRAWN');assert.equal(f.audit.filter(x=>x.state==='WITHDRAW_INTENT').length,1);});
test('withdrawal rejects a different destination',()=>{const f=fixture(),r=f.register();f.store.id=()=> 'different';assert.throws(()=>f.run('withdraw',{operationId:r.operationId}),/another destination/);assert.equal(f.rows[0].withdrawn,false);});
test('history restores operation controls and original link after reload or an uncertain result',()=>{
 const f=fixture(),r=f.register();f.context.taskUrl='';const h=f.run('history');assert.equal(h.state,'VERIFIED');assert.equal(h.operationId,r.operationId);assert.match(h.taskUrl,/videoSet=pages/);
 f.state.failWithdraw=true;f.run('withdraw',{operationId:r.operationId});assert.equal(f.run('history').state,'RECONCILIATION_REQUIRED');
});
