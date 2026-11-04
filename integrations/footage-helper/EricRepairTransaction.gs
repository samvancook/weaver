// Spreadsheet repair inputs belong only to Form Responses 1. Column E is a
// generated filename output, populated strictly from validated source O/AW.
function ericInvariantPatches_(plan){return plan.patches.concat(plan.donte.map(function(t){return {generated:true,row:t.titleRow};}));}
function ericFormSnapshot_(form,ids){
 ids=ids||['1479705159','1675652896','1653265853'];
 return ids.map(function(id){const item=form.getItemById(Number(id)),type=item.getType(),s={id:Number(id),title:item.getTitle(),type:String(type)};
  if(type===FormApp.ItemType.TEXT){s.help=item.asTextItem().getHelpText();return s;}
  const choice=type===FormApp.ItemType.MULTIPLE_CHOICE?item.asMultipleChoiceItem():type===FormApp.ItemType.LIST?item.asListItem():null;if(!choice)throw new Error('Unsupported performer control type.');
  s.values=[];s.navigation={};choice.getChoices().forEach(function(c,i){s.values.push(c.getValue());try{const page=c.getGotoPage();if(page)s.navigation[i]={pageId:page.getId()};else{const nav=c.getPageNavigationType();if(nav)s.navigation[i]={type:String(nav)};}}catch(ignore){}});
  if(type===FormApp.ItemType.MULTIPLE_CHOICE)s.other=choice.hasOtherOption();return s;
 });
}
function ericRestoreFormSnapshot_(form,snapshot){snapshot.forEach(function(s){const item=form.getItemById(s.id);if(String(item.getType())!==s.type||item.getTitle()!==s.title)throw new Error('Form structure changed; restore refused.');if(s.type==='TEXT'){item.asTextItem().setHelpText(s.help);return;}const choice=s.type==='MULTIPLE_CHOICE'?item.asMultipleChoiceItem():item.asListItem();const choices=s.values.map(function(v,i){const nav=s.navigation[i];return nav&&nav.pageId?choice.createChoice(v,form.getItemById(nav.pageId).asPageBreakItem()):nav&&nav.type?choice.createChoice(v,FormApp.PageNavigationType[nav.type]):choice.createChoice(v);});choice.setChoices(choices);if(s.type==='MULTIPLE_CHOICE')choice.showOtherOption(s.other);});}
function ericGeneratedFilenames_(ss,plan){
 const source=ss.getSheetByName('Form Responses 1'),titles=ss.getSheetByName('All titles w/performer & Event');
 plan.donte.forEach(function(t){const name=source.getRange(t.sourceRow,15).getValue(),url=source.getRange(t.sourceRow,49).getValue(),out=titles.getRange(t.titleRow,5);if(name!==t.name||url!==t.url||out.getFormula()||![t.filenameBefore,t.name].includes(out.getValue()))throw new Error('Filename source/output changed; population refused.');out.setValue(name);});SpreadsheetApp.flush();
}
function ericAssertSourcePlan_(plan){
 const allowed=['C30','CH6974','CH6822','L6823','L6837','L6839','L6842','AT6705'];for(let r=6647;r<=6652;r++){allowed.push('AW'+r,'O'+r);}
 if(plan.patches.length!==20||new Set(plan.patches.map(function(p){return p.a1;})).size!==20||plan.patches.some(function(p){return p.sheet!=='Form Responses 1'||allowed.indexOf(p.a1)<0;}))throw new Error('Only the 20 checked source input cells are allowed.');
}
function ericValidateSourceSchema_(ss){
 const sh=ss.getSheetByName('Form Responses 1'),expected={3:'Select Performer',12:'Select Performer',15:'Include the file name including the extension (ex 103.mov)',46:'Rights Management',49:'Copy and paste the link of the file on the drive',86:'If not listed, enter performer here:'},out=[];
 for(const key in expected){const col=Number(key),cell=sh.getRange(1,col),title=cell.getDisplayValue().trim();if(title!==expected[key]||cell.getFormula())throw new Error('Unexpected source schema at column '+col);out.push({column:col,header:title});}return out;
}
function ericExecutePlan_(ss,form,book,plan,options){
 options=options||{};ericAssertSourcePlan_(plan);if(options.fixture?ss.getId()===RETITLE_CONFIG.spreadsheetId:ss.getId()!==RETITLE_CONFIG.spreadsheetId)throw new Error('Unexpected repair destination.');const beforeForm=ericFormSnapshot_(form,options.formIds),id=Utilities.getUuid();
 ericLog_(book,id,'INTENT',{plan:plan,formIds:beforeForm.map(function(s){return s.id;}),formBeforeHashes:beforeForm.map(curationRegistrationHash_)});
 try{
  beforeForm.forEach(function(s){ericLog_(book,id,'FORM_BEFORE',{snapshot:s});});
  plan.patches.forEach(function(p){const cell=ss.getSheetByName('Form Responses 1').getRange(p.a1);if(cell.getFormula()||cell.getValue()!==p.before)throw new Error('Source changed after preview: '+p.a1);});
  plan.patches.forEach(function(p){if(p.before!==p.after)ss.getSheetByName('Form Responses 1').getRange(p.a1).setValue(p.after);});SpreadsheetApp.flush();ericLog_(book,id,'SOURCE_WRITTEN',{sourceCells:plan.patches.length});
  ericGeneratedFilenames_(ss,plan);ericLog_(book,id,'FILENAMES_POPULATED_FROM_SOURCE',{count:plan.donte.length});
  if(JSON.stringify(ericFormSnapshot_(form,options.formIds))!==JSON.stringify(beforeForm))throw new Error('Form changed since backup; refresh refused.');
  const names=ericPerformerNames_(ss),targetForm=beforeForm.map(function(s){const t=JSON.parse(JSON.stringify(s));if(t.type==='TEXT')t.help=names.join('\n');else{t.values=names;t.navigation={};}return t;});
  targetForm.forEach(function(s){ericLog_(book,id,'FORM_TARGET',{snapshot:s});});
  if(options.formIds){targetForm.forEach(function(s){const item=form.getItemById(s.id);if(s.type==='TEXT')item.asTextItem().setHelpText(s.help);else if(s.type==='MULTIPLE_CHOICE')item.asMultipleChoiceItem().setChoiceValues(s.values);else item.asListItem().setChoiceValues(s.values);});}else ericRefreshPickers_(ss,form);
  const afterForm=ericFormSnapshot_(form,options.formIds);afterForm.forEach(function(s){ericLog_(book,id,'FORM_AFTER',{snapshot:s});});
  const result=options.verify?options.verify():ericVerify_(ss,form,plan);if(!result.verified)throw new Error(JSON.stringify(result.errors));
  ericLog_(book,id,'VERIFIED',{result:result});return {state:'VERIFIED',operationId:id,auditUrl:book.getUrl(),result:result};
 }catch(error){ericLog_(book,id,'UNCERTAIN',{error:String(error.message||error)});return {state:'UNCERTAIN',operationId:id,auditUrl:book.getUrl(),error:String(error.message||error)};}
}
function ericUndoPlan_(ss,form,book,operationId){
 const history=ericHistory_(book),own=history.filter(function(h){return h.operationId===operationId;}),intent=own.find(function(h){return h.state==='INTENT';});if(!intent)throw new Error('Unknown owned repair.');if(own[own.length-1].state==='UNDONE')return {state:'UNDONE',operationId:operationId};
 const plan=intent.details.plan;ericAssertSourcePlan_(plan);const before=own.filter(function(h){return h.state==='FORM_BEFORE';}).map(function(h){return h.details.snapshot;}),target=own.filter(function(h){return h.state==='FORM_TARGET';}).map(function(h){return h.details.snapshot;}),after=own.filter(function(h){return h.state==='FORM_AFTER';}).map(function(h){return h.details.snapshot;});
 if(before.length!==intent.details.formIds.length)throw new Error('Complete form backup is missing; no undo writes attempted.');
 if(JSON.stringify(ericSourceInvariant_(ss,ericInvariantPatches_(plan)))!==JSON.stringify(plan.beforeInvariant))throw new Error('Unrelated source state changed; undo refused.');
 plan.patches.forEach(function(p){const cell=ss.getSheetByName(p.sheet).getRange(p.a1);if(cell.getFormula()||![p.before,p.after].includes(cell.getValue()))throw new Error('A repair input changed; undo refused.');});
 plan.donte.forEach(function(t){const cell=ss.getSheetByName('All titles w/performer & Event').getRange(t.titleRow,5);if(cell.getFormula()||![t.filenameBefore,t.name].includes(cell.getValue()))throw new Error('A generated filename changed; undo refused.');});
 const current=ericFormSnapshot_(form,intent.details.formIds);current.forEach(function(s,i){if(![before[i],target[i],after[i]].filter(Boolean).some(function(v){return JSON.stringify(v)===JSON.stringify(s);}))throw new Error('Form setting changed; undo refused.');});
 ericLog_(book,operationId,'UNDO_INTENT',{sourceCells:plan.patches.length,generatedFilenames:plan.donte.length});
 try{
  plan.patches.forEach(function(p){ss.getSheetByName('Form Responses 1').getRange(p.a1).setValue(p.before);});SpreadsheetApp.flush();
  plan.donte.forEach(function(t){ss.getSheetByName('All titles w/performer & Event').getRange(t.titleRow,5).setValue(t.filenameBefore);});ericRestoreFormSnapshot_(form,before);SpreadsheetApp.flush();
  plan.patches.forEach(function(p){if(ss.getSheetByName(p.sheet).getRange(p.a1).getValue()!==p.before)throw new Error('Source undo mismatch.');});plan.donte.forEach(function(t){if(ss.getSheetByName('All titles w/performer & Event').getRange(t.titleRow,5).getValue()!==t.filenameBefore)throw new Error('Generated filename undo mismatch.');});
  if(JSON.stringify(ericFormSnapshot_(form,intent.details.formIds))!==JSON.stringify(before)||JSON.stringify(ericSourceInvariant_(ss,ericInvariantPatches_(plan)))!==JSON.stringify(plan.beforeInvariant))throw new Error('Undo form/invariant readback mismatch.');
  ericLog_(book,operationId,'UNDONE',{sourceInputsRestored:true,generatedFilenamesRestored:true,exactFormSettingsRestored:true});return {state:'UNDONE',operationId:operationId,auditUrl:book.getUrl()};
 }catch(error){ericLog_(book,operationId,'UNCERTAIN_UNDO',{error:String(error.message||error)});throw error;}
}
