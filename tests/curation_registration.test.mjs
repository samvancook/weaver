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
test('a picker failure is uncertain; explicit derived recovery never appends the registration twice',()=>{
 const f=fixture();let pickerMatches=false,refreshFails=true,refreshCount=0;
 const verifySource=f.store.verify;
 f.store.sourceState=op=>f.rows.some(r=>r.op.operationId===op.operationId&&!r.changed)?f.rows[0].withdrawn?'withdrawn':'active':'changed';
 f.store.verify=(op,withdrawn)=>verifySource(op,withdrawn)&&pickerMatches;
 f.store.refreshDerived=()=>{refreshCount++;if(refreshFails)throw Error('picker unavailable');pickerMatches=true;return {closePicker:['PAGES']};};
 const r=f.register();assert.equal(r.state,'RECONCILIATION_REQUIRED');assert.equal(f.run('reconcile',{operationId:r.operationId}).state,'RECONCILIATION_REQUIRED');
 assert.equal(f.register().state,'RECONCILIATION_REQUIRED');assert.equal(f.rows.length,1);
 refreshFails=false;assert.equal(f.run('repairDerived',{operationId:r.operationId}).state,'VERIFIED');assert.equal(f.rows.length,1);assert.equal(refreshCount,2);
 assert.ok(f.audit.some(x=>x.state==='DERIVED_REPAIR_INTENT'));
 f.rows[0].changed=true;assert.equal(f.run('repairDerived',{operationId:r.operationId}).state,'CHANGED_RECORD_PRESERVED');assert.equal(refreshCount,2);
});
test('a crash after derived update remains unresolved and recoverable from history',()=>{
 const f=fixture();f.store.refreshDerived=()=>({closePicker:['PAGES']});f.state.failAudit='VERIFIED';const r=f.register();assert.equal(r.state,'RECONCILIATION_REQUIRED');assert.equal(f.run('history').state,'RECONCILIATION_REQUIRED');assert.equal(f.register().state,'RECONCILIATION_REQUIRED');assert.equal(f.rows.length,1);
});
const integrationScope={};vm.createContext(integrationScope);vm.runInContext(fs.readFileSync(new URL('../integrations/footage-helper/CurationRegistrationIntegration.gs',import.meta.url),'utf8'),integrationScope);
test('derived verifier requires exact open-event changes, picker parity, retained other links and unchanged closure',()=>{
 const before={curationLinks:[['legacy','old']],needingCuration:['open','closed'],eventsNeedingCuration:['open','closed'],readyToClose:['legacy'],closed:['closed'],closePicker:['legacy']};
 const after={curationLinks:[['legacy','old'],['open','new']],needingCuration:['closed'],eventsNeedingCuration:['closed'],readyToClose:['legacy','open'],closed:['closed'],closePicker:['legacy','open']};
 const verify=integrationScope.curationIntegrationVerifyDerived_;
 assert.equal(verify(before,after,'open','new',false),true);
 for(const key of ['curationLinks','needingCuration','eventsNeedingCuration','readyToClose','closed','closePicker']){const bad=structuredClone(after);bad[key]=[];assert.equal(verify(before,bad,'open','new',false),false,key);}
 assert.equal(verify(before,before,'open','new',true),true);
});
test('a registered closed event never becomes available to close; empty picker uses canonical placeholder',()=>{
 const before={curationLinks:[],needingCuration:['closed'],eventsNeedingCuration:['closed'],readyToClose:[],closed:['closed'],closePicker:['— No events to close —']};
 const after={...before,curationLinks:[['closed','new']],needingCuration:[],eventsNeedingCuration:[]};
 assert.equal(integrationScope.curationIntegrationVerifyDerived_(before,after,'closed','new',false),true);
 const bad=structuredClone(after);bad.readyToClose=['closed'];bad.closePicker=['closed'];assert.equal(integrationScope.curationIntegrationVerifyDerived_(before,bad,'closed','new',false),false);
});
test('live preflight reads current receipt, closure and picker without exposing a production write',()=>{
 const state={ownerChecks:0,links:[],closed:[],picker:['— No events to close —']};
 const live={...integrationScope,weaverRequireOwner_:()=>state.ownerChecks++,curationRegistrationSourceContext_:()=>({eventName:'PAGES',taskUrl:'https://weaver.test/task',sourceEventId:'receipt',prioritySetId:'pages',sourceSnapshot:{closed:false,existingCurationLinks:state.links}}),
 getRetitleSpreadsheet_:()=>({getId:()=> 'production-id'}),Session:{getActiveUser:()=>({getEmail:()=> 'sam@test'})},
 FormApp:{ItemType:{MULTIPLE_CHOICE:'MC'},openById:()=>({getItemById:()=>({getType:()=> 'MC',asMultipleChoiceItem:()=>({getChoices:()=>state.picker.map(value=>({getValue:()=>value}))})})})}};
 vm.createContext(live);vm.runInContext(fs.readFileSync(new URL('../integrations/footage-helper/CurationRegistrationIntegration.gs',import.meta.url),'utf8'),live);
 live.curationIntegrationColumn_=(_book,sheet,col)=>col===9?state.closed:col===10?[]:['PAGES'];
 const r=live.previewInventoryCurationRegistration('PAGES');assert.equal(r.productionEnabled,false);assert.equal(r.canRegister,false);assert.equal(r.blockers.length,0);assert.equal(r.plannedEffect.addToClosePicker,true);assert.equal(state.ownerChecks,1);
 state.picker=['stale'];assert.match(live.previewInventoryCurationRegistration('PAGES').blockers.join(' '),/picker differs/);
 state.links=[{url:'legacy'}];assert.match(live.previewInventoryCurationRegistration('PAGES').blockers.join(' '),/existing curation link/);
});
