// One-event production adapter. Disabled until the deploying owner authorizes
// the exact current preview. No permission, formula, media or closure changes.
const CURATION_PILOT_EVENT='PAGES footage from Rudy Francisco 2026';
const CURATION_PILOT_AUDIT_PROP='WEAVER_CURATION_PILOT_AUDIT';
const CURATION_PILOT_APPROVAL_PROP='WEAVER_CURATION_PILOT_APPROVAL';
const CURATION_INVENTORY_FORM_ID='1Tmr8J24JMtAmGS2T0eKyKytnLvlx778CM7Hu8g5MP80';
const CURATION_INVENTORY_CLOSE_ITEM=466810607;
function curationProductionGuard_(eventName,requireApproval) {
  weaverRequireOwner_();
  if(eventName!==CURATION_PILOT_EVENT)throw new Error('The inventory pilot is limited to the exact PAGES event.');
  if(requireApproval){const approved=PropertiesService.getScriptProperties().getProperty(CURATION_PILOT_APPROVAL_PROP);if(!approved)throw new Error('The production pilot is not authorized.');}
}
function curationPilotAuditBook_() {
  const props=PropertiesService.getScriptProperties(),id=props.getProperty(CURATION_PILOT_AUDIT_PROP);
  if(id){if(id===RETITLE_CONFIG.spreadsheetId)throw new Error('The audit must be in a Helper-owned file.');return SpreadsheetApp.openById(id);}
  const book=SpreadsheetApp.create('Footage Helper — PAGES curation registration audit');
  const sheet=book.getSheets()[0];sheet.setName('Audit');sheet.appendRow(CURATION_REGISTRATION_AUDIT_HEADERS);sheet.setFrozenRows(1);
  props.setProperty(CURATION_PILOT_AUDIT_PROP,book.getId());return book;
}
function curationPilotChoiceItem_(form) {
  const item=form.getItemById(CURATION_INVENTORY_CLOSE_ITEM);if(!item)throw new Error('Live close-event picker is missing.');
  switch(item.getType()){
    case FormApp.ItemType.MULTIPLE_CHOICE:return item.asMultipleChoiceItem();
    case FormApp.ItemType.LIST:return item.asListItem();
    case FormApp.ItemType.CHECKBOX:return item.asCheckboxItem();
    default:throw new Error('Live close-event picker type changed.');
  }
}
// Hash unrelated event data instead of copying the full inventory into the audit.
function curationPilotDerivedSnapshot_(book,form,eventName) {
  const raw=curationIntegrationSnapshot_(book,form,CURATION_INVENTORY_CLOSE_ITEM);
  function withoutEvent(values){return values.filter(function(v){return v!==eventName&&v!==CURATION_EMPTY_CLOSE_CHOICE;}).sort();}
  return {event:{curationLinks:raw.curationLinks.filter(function(v){return v[0]===eventName;}).map(function(v){return v[1];}),needs:raw.needingCuration.indexOf(eventName)>=0,eventNeeds:raw.eventsNeedingCuration.indexOf(eventName)>=0,ready:raw.readyToClose.indexOf(eventName)>=0,closed:raw.closed.indexOf(eventName)>=0,picker:raw.closePicker.indexOf(eventName)>=0},
    otherHashes:{curation:curationRegistrationHash_(raw.curationLinks.filter(function(v){return v[0]!==eventName;}).sort()),needs:curationRegistrationHash_(withoutEvent(raw.needingCuration)),eventNeeds:curationRegistrationHash_(withoutEvent(raw.eventsNeedingCuration)),ready:curationRegistrationHash_(withoutEvent(raw.readyToClose)),closed:curationRegistrationHash_(raw.closed.slice().sort()),picker:curationRegistrationHash_(withoutEvent(raw.closePicker))},
    pickerMatches:JSON.stringify(raw.closePicker.slice().sort())===JSON.stringify((raw.readyToClose.length?raw.readyToClose:[CURATION_EMPTY_CLOSE_CHOICE]).slice().sort())};
}
function curationPilotVerifyDerived_(before,after,url,withdrawn) {
  const expected=withdrawn?before.event:{curationLinks:[url],needs:false,eventNeeds:false,ready:true,closed:false,picker:true};
  return after.pickerMatches&&JSON.stringify(after.otherHashes)===JSON.stringify(before.otherHashes)&&JSON.stringify(after.event)===JSON.stringify(expected);
}
function curationPilotStore_(action) {
  curationProductionGuard_(CURATION_PILOT_EVENT,false);
  const book=getRetitleSpreadsheet_();if(book.getId()!==RETITLE_CONFIG.spreadsheetId)throw new Error('Inventory destination changed.');
  const form=FormApp.openById(CURATION_INVENTORY_FORM_ID),auditBook=curationPilotAuditBook_();
  // No current receipt is required to inspect or reverse an already-owned operation.
  const context=(action==='preview'||action==='register')?curationRegistrationSourceContext_(CURATION_PILOT_EVENT):{eventName:CURATION_PILOT_EVENT,taskUrl:'',sourceSnapshot:{existingCurationLinks:[]}};
  if(action==='preview'||action==='register'){
    context.sourceSnapshot.derivedBefore=curationPilotDerivedSnapshot_(book,form,CURATION_PILOT_EVENT);
    const d=context.sourceSnapshot.derivedBefore;
    if(!d.pickerMatches||d.event.closed||!d.event.needs||!d.event.eventNeeds) {
      // Existing registrations are returned intact, including after a successful retry.
      if(!context.sourceSnapshot.existingCurationLinks.length)throw new Error('The PAGES pilot baseline changed; no source write is allowed.');
    }
  }
  const store=curationRegistrationSpreadsheetStore_(book,context,{pilot:true,auditBook:auditBook,action:action}),verifySourceAndLink=store.verify;
  store.authorization=function(){return JSON.parse(PropertiesService.getScriptProperties().getProperty(CURATION_PILOT_APPROVAL_PROP)||'null');};
  store.destination=function(){return 'inventory';};store.url=function(){return auditBook.getUrl();};store.auditUrl=store.url;
  store.refreshDerived=function(op,withdrawn){
    curationProductionGuard_(op.eventName,false);
    if(store.sourceState(op)!==(withdrawn?'withdrawn':'active'))throw new Error('Owned registration changed; picker was preserved.');
    const current=curationPilotDerivedSnapshot_(book,form,op.eventName),before=op.sourceSnapshot.derivedBefore;
    // Do not silently repair unrelated changes that happened after this preview.
    for(const key in before.otherHashes)if(before.otherHashes[key]!==current.otherHashes[key])throw new Error('Unrelated inventory state changed; preserve the picker for review.');
    const names=curationIntegrationColumn_(book,'Event Worksheet',10),choices=names.length?names:[CURATION_EMPTY_CLOSE_CHOICE];
    curationPilotChoiceItem_(form).setChoiceValues(choices);
    const after=curationPilotDerivedSnapshot_(book,form,op.eventName);
    if(!curationPilotVerifyDerived_(before,after,op.taskUrl,withdrawn))throw new Error('Inventory/picker readback did not verify.');
    return after;
  };
  store.verify=function(op,withdrawn){return verifySourceAndLink(op,withdrawn)&&curationPilotVerifyDerived_(op.sourceSnapshot.derivedBefore,curationPilotDerivedSnapshot_(book,form,op.eventName),op.taskUrl,withdrawn);};
  return store;
}
function curationPilotRun_(action,eventName,extra) {
  curationProductionGuard_(eventName,action==='register');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const store=curationPilotStore_(action),input=Object.assign({eventName:eventName},extra||{});
    if(action==='register'){
      const approval=JSON.parse(PropertiesService.getScriptProperties().getProperty(CURATION_PILOT_APPROVAL_PROP)||'null');
      if(!approval||approval.previewKey!==input.previewKey)throw new Error('The exact preview has not been authorized.');
      if(store.history(eventName).some(function(r){return r.state==='INTENT';})&&!store.active(eventName).length)throw new Error('This one-registration pilot has already been used. Inspect its history; no second append is allowed.');
    }
    const result=curationRegistrationEngine_(store,action,input);
    result.productionEnabled=!!PropertiesService.getScriptProperties().getProperty(CURATION_PILOT_APPROVAL_PROP);
    result.pilotEvent=CURATION_PILOT_EVENT;
    if(action==='preview'&&!result.productionEnabled){result.canRegister=false;result.message='PAGES pilot prepared, but live recording is disabled pending authorization. Production inventory was not changed.';}
    return result;
  }finally{lock.releaseLock();}
}
function prepareInventoryCurationRegistration(eventName){return curationPilotRun_('preview',eventName);}
// This owner-only capability is invoked only after approval of the exact live pilot.
function authorizeInventoryCurationPilot(eventName,previewKey) {
  curationProductionGuard_(eventName,false);const preview=prepareInventoryCurationRegistration(eventName);
  if(preview.state!=='PREVIEW'||preview.previewKey!==previewKey)throw new Error('The current PAGES preview changed.');
  PropertiesService.getScriptProperties().setProperty(CURATION_PILOT_APPROVAL_PROP,JSON.stringify({eventName:eventName,previewKey:previewKey,actor:Session.getActiveUser().getEmail(),at:new Date().toISOString()}));
  return {state:'PILOT_AUTHORIZED',eventName:eventName};
}
function registerInventoryCurationLink(eventName,previewKey){return curationPilotRun_('register',eventName,{previewKey:previewKey});}
function withdrawInventoryCurationLink(eventName,operationId){return curationPilotRun_('withdraw',eventName,{operationId:operationId});}
function repairInventoryCurationDerivedState(eventName,operationId){return curationPilotRun_('repairDerived',eventName,{operationId:operationId});}
function getInventoryCurationRegistrationHistory(eventName){return curationPilotRun_('history',eventName);}
