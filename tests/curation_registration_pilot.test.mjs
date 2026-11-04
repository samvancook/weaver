import test from 'node:test';import assert from 'node:assert/strict';import fs from 'node:fs';import vm from 'node:vm';
const file=fs.readFileSync(new URL('../integrations/footage-helper/CurationRegistrationPilot.gs',import.meta.url),'utf8');
function scope(){let approval='',ownerChecks=0;const s={weaverRequireOwner_:()=>ownerChecks++,PropertiesService:{getScriptProperties:()=>({getProperty:()=>approval})}};vm.createContext(s);vm.runInContext(file,s);return {s,approve:()=>approval='approved',checks:()=>ownerChecks};}
test('pilot rejects every other event and defaults to disabled for production appends',()=>{const f=scope();assert.throws(()=>f.s.curationProductionGuard_('MPMU',false),/exact PAGES/);assert.throws(()=>f.s.curationProductionGuard_('PAGES footage from Rudy Francisco 2026',true),/not authorized/);f.s.curationProductionGuard_('PAGES footage from Rudy Francisco 2026',false);f.approve();f.s.curationProductionGuard_('PAGES footage from Rudy Francisco 2026',true);assert.equal(f.checks(),4);});
test('production verifier preserves unrelated hashes and requires exact restoration',()=>{
 const {s}=scope(),before={event:{curationLinks:[],needs:true,eventNeeds:true,ready:false,closed:false,picker:false},otherHashes:{curation:'a',needs:'b',eventNeeds:'c',ready:'d',closed:'e',picker:'f'},pickerMatches:true};
 const after={event:{curationLinks:['https://weaver/task'],needs:false,eventNeeds:false,ready:true,closed:false,picker:true},otherHashes:{...before.otherHashes},pickerMatches:true};
 assert.equal(s.curationPilotVerifyDerived_(before,after,'https://weaver/task',false),true);
 for(const key of Object.keys(before.otherHashes)){const bad=structuredClone(after);bad.otherHashes[key]='changed';assert.equal(s.curationPilotVerifyDerived_(before,bad,'https://weaver/task',false),false,key);}
 assert.equal(s.curationPilotVerifyDerived_(before,before,'https://weaver/task',true),true);
 const changed=structuredClone(before);changed.event.closed=true;assert.equal(s.curationPilotVerifyDerived_(before,changed,'https://weaver/task',true),false);
});
test('shared spreadsheet adapter still refuses arbitrary production use and guards before source append',()=>{
 let writes=0,guardCalls=0;const source={getLastColumn:()=>99,appendRow:()=>writes++};const audit={};const book={getId:()=> 'production',getSheetByName:name=>name==='Form Responses 1'?source:name==='Curation'?{}:null};
 const s={RETITLE_CONFIG:{spreadsheetId:'production'},curationProductionGuard_:(_event,write)=>{guardCalls++;if(write)throw Error('disabled');}};vm.createContext(s);vm.runInContext(fs.readFileSync(new URL('../integrations/footage-helper/CurationRegistration.gs',import.meta.url),'utf8'),s);
 assert.throws(()=>s.curationRegistrationSpreadsheetStore_(book,{eventName:'PAGES'}),/not an allowed destination/);
 assert.throws(()=>s.curationRegistrationSpreadsheetStore_(book,{eventName:'PAGES'},{pilot:true,auditBook:book}),/not an allowed destination/);
 const ownAudit={getId:()=> 'own-audit',getSheetByName:()=>audit};const store=s.curationRegistrationSpreadsheetStore_(book,{eventName:'PAGES'},{pilot:true,auditBook:ownAudit,action:'register'});
 assert.throws(()=>store.append({}),/disabled/);assert.equal(writes,0);assert.equal(guardCalls,2);
});
