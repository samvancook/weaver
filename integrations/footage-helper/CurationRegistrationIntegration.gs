// Synthetic-only integration harness. Copies no production inventory records.
const CURATION_INTEGRATION_PROP='WEAVER_CURATION_INTEGRATION_FIXTURE';
const CURATION_TEST_EVENTS=['TEST ONLY — open event','TEST ONLY — closed event','TEST ONLY — existing curation event'];
const CURATION_EMPTY_CLOSE_CHOICE='— No events to close —';

function curationIntegrationSetup_() {
  const props=PropertiesService.getScriptProperties(),existing=props.getProperty(CURATION_INTEGRATION_PROP);
  if(existing)return JSON.parse(existing);
  const book=SpreadsheetApp.create('Footage Helper — synthetic inventory integration test');
  const source=book.getSheets()[0];source.setName('Form Responses 1');
  source.insertColumnsAfter(source.getMaxColumns(),99-source.getMaxColumns());
  const header=new Array(99).fill('');header[0]='Timestamp';header[1]='Select an option';header[23]='Event';header[47]='Email';header[49]='Curation event';header[50]='Curation link';header[78]='Closed event';source.appendRow(header);
  CURATION_TEST_EVENTS.forEach(function(name){const row=new Array(99).fill('');row[0]=new Date();row[1]='Add Event';row[23]=name;source.appendRow(row);});
  let row=new Array(99).fill('');row[0]=new Date();row[1]='Close Event';row[78]=CURATION_TEST_EVENTS[1];source.appendRow(row);
  row=new Array(99).fill('');row[0]=new Date();row[1]='Add Curation Form';row[49]=CURATION_TEST_EVENTS[2];row[50]='https://example.invalid/legacy-curation';source.appendRow(row);
  const view=book.insertSheet('Curation');view.getRange('A1:D1').setValues([['Curation Form','Curation Form Link','','Unique events']]);
  view.getRange('A2').setFormula('=IFERROR(FILTER(\'Form Responses 1\'!AX2:AX,\'Form Responses 1\'!B2:B="Add Curation Form"),"")');
  view.getRange('B2').setFormula('=IFERROR(FILTER(\'Form Responses 1\'!AY2:AY,\'Form Responses 1\'!B2:B="Add Curation Form"),"")');
  view.getRange('D2').setFormula('=UNIQUE(A2:A)');
  const events=book.insertSheet('Events');
  events.getRange('A1').setFormula('={"Event Name","Event Date","Venue Name";IFERROR(ARRAYFORMULA(TRIM(CLEAN(FILTER(\'Form Responses 1\'!X2:Z,\'Form Responses 1\'!B2:B="Add Event")))),"")}');
  events.getRange('O1:S1').setValues([['Unique events','','With curation link','Unique with link','Needs curation link']]);
  events.getRange('O2').setFormula('=UNIQUE(A2:A)');events.getRange('Q2').setFormula('=IFERROR(FILTER(\'Form Responses 1\'!AX2:AX,\'Form Responses 1\'!B2:B="Add Curation Form"),"")');
  events.getRange('R2').setFormula('=UNIQUE(Q2:Q)');events.getRange('S2').setFormula('=IFERROR(FILTER(O2:O,LEN(O2:O),ISERROR(MATCH(O2:O,Q2:Q,0))),"")');
  const titles=book.insertSheet('All titles w/performer & Event');titles.getRange('D1:D4').setValues([['Synthetic event names'],[CURATION_TEST_EVENTS[0]],[CURATION_TEST_EVENTS[1]],[CURATION_TEST_EVENTS[2]]]);
  const worksheet=book.insertSheet('Event Worksheet');worksheet.getRange('A1:J1').setValues([['Events','With curation','Unique with curation','Needs curation','Titles','Without titles','','','Closed','Events to Close']]);
  worksheet.getRange('A2').setFormula('=UNIQUE(Events!A2:A)');
  worksheet.getRange('B1').setFormula('={"Events With Curation Forms";IFERROR(ARRAYFORMULA(TRIM(CLEAN(FILTER(\'Form Responses 1\'!AX2:AX,\'Form Responses 1\'!B2:B="Add Curation Form")))),"")}');
  worksheet.getRange('C2').setFormula('=UNIQUE(B2:B)');worksheet.getRange('D2').setFormula('=IFERROR(FILTER(A2:A,LEN(A2:A),ISERROR(MATCH(A2:A,B2:B,0))),"")');
  worksheet.getRange('E2').setFormula('=UNIQUE(\'All titles w/performer & Event\'!D2:D)');worksheet.getRange('F2').setFormula('=IFERROR(FILTER(A2:A,LEN(A2:A),ISERROR(MATCH(A2:A,E2:E,0))),"")');
  worksheet.getRange('I2').setFormula('=IFERROR(FILTER(\'Form Responses 1\'!CA2:CA,\'Form Responses 1\'!B2:B="Close Event"),"")');
  worksheet.getRange('J2').setFormula('=IFERROR(FILTER(C2:C,LEN(C2:C),ISERROR(MATCH(C2:C,I2:I,0))),"")');
  book.insertSheet('Audit').appendRow(CURATION_REGISTRATION_AUDIT_HEADERS);
  const form=FormApp.create('Footage Helper — synthetic close-picker test');form.setAcceptingResponses(false);
  const item=form.addMultipleChoiceItem().setTitle('TEST ONLY — Event to close');
  const config={spreadsheetId:book.getId(),formId:form.getId(),closeItemId:item.getId(),createdAt:new Date().toISOString(),syntheticOnly:true};
  curationIntegrationRefreshPicker_(book,form,item.getId());
  props.setProperty(CURATION_INTEGRATION_PROP,JSON.stringify(config));return config;
}
function curationIntegrationColumn_(book,sheet,column) {
  const sh=book.getSheetByName(sheet),last=sh.getLastRow();
  const values=last<2?[]:sh.getRange(2,column,last-1,1).getDisplayValues().map(function(v){return v[0];}).filter(Boolean);
  if(values.some(function(v){return /^#(REF!|ERROR!|N\/A|VALUE!|NAME\?|DIV\/0!|NUM!)/.test(v);}))throw new Error('Derived formula error in '+sheet+' column '+column);
  return values;
}
function curationIntegrationRefreshPicker_(book,form,itemId) {
  SpreadsheetApp.flush();
  const names=curationIntegrationColumn_(book,'Event Worksheet',10);
  const item=form.getItemById(itemId);
  if(!item||item.getType()!==FormApp.ItemType.MULTIPLE_CHOICE)throw new Error('The owned test close picker is missing or changed.');
  const choices=names.length?names:[CURATION_EMPTY_CLOSE_CHOICE];item.asMultipleChoiceItem().setChoiceValues(choices);
  const actual=item.asMultipleChoiceItem().getChoices().map(function(c){return c.getValue();});
  if(JSON.stringify(actual)!==JSON.stringify(choices))throw new Error('Close picker readback failed.');
  return {readyToClose:names,closePicker:actual};
}
function curationIntegrationSnapshot_(book,form,itemId) {
  SpreadsheetApp.flush();
  const view=book.getSheetByName('Curation'),n=view.getLastRow();
  return {curationLinks:n<2?[]:view.getRange(2,1,n-1,2).getDisplayValues().filter(function(v){return !!v[0];}),
    needingCuration:curationIntegrationColumn_(book,'Event Worksheet',4),
    eventsNeedingCuration:curationIntegrationColumn_(book,'Events',19),
    readyToClose:curationIntegrationColumn_(book,'Event Worksheet',10),
    closed:curationIntegrationColumn_(book,'Event Worksheet',9),
    closePicker:curationRegistrationChoiceValues_(form,itemId)};
}
function curationIntegrationVerifyDerived_(before,after,eventName,taskUrl,withdrawn) {
  function sameSet(a,b){return JSON.stringify(a.slice().sort())===JSON.stringify(b.slice().sort());}
  const needs=withdrawn?before.needingCuration:before.needingCuration.filter(function(v){return v!==eventName;});
  const eventNeeds=withdrawn?before.eventsNeedingCuration:before.eventsNeedingCuration.filter(function(v){return v!==eventName;});
  const ready=withdrawn||before.closed.indexOf(eventName)>=0?before.readyToClose:before.readyToClose.concat([eventName]);
  const links=withdrawn?before.curationLinks:before.curationLinks.concat([[eventName,taskUrl]]);
  return sameSet(after.needingCuration,needs)&&sameSet(after.eventsNeedingCuration,eventNeeds)&&sameSet(after.readyToClose,ready)&&sameSet(after.closed,before.closed)&&
    sameSet(after.curationLinks.map(function(v){return JSON.stringify(v);}),links.map(function(v){return JSON.stringify(v);}))&&sameSet(after.closePicker,ready.length?ready:[CURATION_EMPTY_CLOSE_CHOICE]);
}
function curationIntegrationStore_(config,eventIndex) {
  if(!config.syntheticOnly||config.spreadsheetId===RETITLE_CONFIG.spreadsheetId||config.formId==='1Tmr8J24JMtAmGS2T0eKyKytnLvlx778CM7Hu8g5MP80')throw new Error('Only the synthetic fixture is allowed.');
  const eventName=CURATION_TEST_EVENTS[eventIndex];if(!eventName)throw new Error('Unknown synthetic event.');
  const book=SpreadsheetApp.openById(config.spreadsheetId),form=FormApp.openById(config.formId),before=curationIntegrationSnapshot_(book,form,config.closeItemId);
  const context={eventName:eventName,taskUrl:'https://example.invalid/weaver?videoSet=synthetic-'+eventIndex,sourceEventId:'synthetic-receipt-'+eventIndex,prioritySetId:'synthetic-'+eventIndex,
    sourceSnapshot:{syntheticOnly:true,existingCurationLinks:before.curationLinks.filter(function(v){return v[0]===eventName;}).map(function(v){return {eventName:v[0],url:v[1]};}),derivedBefore:before}};
  const store=curationRegistrationSpreadsheetStore_(book,context),verifySourceAndLink=store.verify;
  store.refreshDerived=function(){return curationIntegrationRefreshPicker_(book,form,config.closeItemId);};
  store.verify=function(op,withdrawn){return verifySourceAndLink(op,withdrawn)&&curationIntegrationVerifyDerived_(op.sourceSnapshot.derivedBefore,curationIntegrationSnapshot_(book,form,config.closeItemId),op.eventName,op.taskUrl,withdrawn);};
  store.snapshot=function(){return curationIntegrationSnapshot_(book,form,config.closeItemId);};return store;
}
// Fixed synthetic events only. No caller-supplied workbook, form, URL or source rows.
function testCurationRegistrationIntegration(input) {
  weaverRequireOwner_();const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    input=input||{};const config=curationIntegrationSetup_(),store=curationIntegrationStore_(config,Number(input.eventIndex||0));
    if(input.simulatePickerFailure)store.refreshDerived=function(){throw new Error('TEST ONLY — simulated close-picker failure before the form setter.');};
    const result=curationRegistrationEngine_(store,input.action||'history',{eventName:CURATION_TEST_EVENTS[Number(input.eventIndex||0)],previewKey:input.previewKey,operationId:input.operationId});
    result.syntheticOnly=true;result.formUrl=FormApp.openById(config.formId).getEditUrl();result.snapshot=store.snapshot();return result;
  }finally{lock.releaseLock();}
}

// Read-only production preflight: no source append, formula edit or form setter.
function previewInventoryCurationRegistration(eventName) {
  weaverRequireOwner_();
  const context=curationRegistrationSourceContext_(eventName),book=getRetitleSpreadsheet_();
  const form=FormApp.openById('1Tmr8J24JMtAmGS2T0eKyKytnLvlx778CM7Hu8g5MP80'),item=form.getItemById(466810607);
  if(!item)throw new Error('The live close-event picker could not be found.');
  const type=item.getType();
  const choiceItem=type===FormApp.ItemType.MULTIPLE_CHOICE?item.asMultipleChoiceItem():type===FormApp.ItemType.LIST?item.asListItem():type===FormApp.ItemType.CHECKBOX?item.asCheckboxItem():null;
  if(!choiceItem)throw new Error('The live close-event picker has an unsupported type.');
  const ready=curationIntegrationColumn_(book,'Event Worksheet',10),picker=choiceItem.getChoices().map(function(c){return c.getValue();});
  const needs=curationIntegrationColumn_(book,'Event Worksheet',4),eventNeeds=curationIntegrationColumn_(book,'Events',19),closed=curationIntegrationColumn_(book,'Event Worksheet',9);
  const expectedPicker=ready.length?ready:[CURATION_EMPTY_CLOSE_CHOICE];
  const pickerMatches=JSON.stringify(picker.slice().sort())===JSON.stringify(expectedPicker.slice().sort());
  const isClosed=closed.indexOf(context.eventName)>=0;
  const blockers=[];
  if(context.sourceSnapshot.existingCurationLinks.length)blockers.push('An existing curation link must be preserved.');
  if(!pickerMatches)blockers.push('The current close-event picker differs from its worksheet source; it needs a separate review.');
  if(isClosed!==!!context.sourceSnapshot.closed)blockers.push('The inventory closure views disagree for this event.');
  if(!needs.includes(context.eventName)||!eventNeeds.includes(context.eventName))blockers.push('The two needs-curation lists do not both contain this exact event.');
  return {state:'PRODUCTION_PREVIEW_ONLY',eventName:context.eventName,taskUrl:context.taskUrl,productionEnabled:false,canRegister:false,history:[],
    message:'Read-only live preview. No inventory rows, formulas or form choices were changed.',
    previewSummary:'Existing curation links: '+context.sourceSnapshot.existingCurationLinks.length+'. Event is '+(isClosed?'closed':'open')+'. Close-picker parity: '+(pickerMatches?'matches':'needs review')+'. '+(blockers.length?blockers.join(' '):'The planned changes are consistent with the current source.'),
    blockers:blockers,sourceEventId:context.sourceEventId,prioritySetId:context.prioritySetId,
    plannedSource:{spreadsheetId:book.getId(),sheet:'Form Responses 1',action:'Add Curation Form',email:Session.getActiveUser().getEmail(),event:context.eventName,url:context.taskUrl,columns:['A','B','AV','AX','AY']},
    plannedEffect:{removeFromNeedsCuration:true,addToClosePicker:!isClosed,preserveClosedState:true,refreshOnlyCloseItemId:466810607},
    current:{closed:isClosed,needsCuration:needs.includes(context.eventName),eventsNeedsCuration:eventNeeds.includes(context.eventName),readyToClose:ready.includes(context.eventName),closePickerMatchesSource:pickerMatches,readyEventCount:ready.length,closePickerChoiceCount:picker.length}};
}

function curationRegistrationChoiceValues_(form,itemId) {
  const item=form.getItemById(itemId);if(!item)throw new Error('Close picker is missing.');
  const type=item.getType(),choiceItem=type===FormApp.ItemType.MULTIPLE_CHOICE?item.asMultipleChoiceItem():type===FormApp.ItemType.LIST?item.asListItem():type===FormApp.ItemType.CHECKBOX?item.asCheckboxItem():null;
  if(!choiceItem)throw new Error('Close picker type changed.');
  return choiceItem.getChoices().map(function(c){return c.getValue();});
}
