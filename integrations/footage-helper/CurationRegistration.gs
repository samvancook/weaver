// First implementation: tool-owned sandbox only. No production response writer.
const CURATION_REGISTRATION_SANDBOX_PROP = 'WEAVER_CURATION_REGISTRATION_SANDBOX';
const CURATION_REGISTRATION_AUDIT_HEADERS = ['Operation ID','Event','Action','State','Actor','Time','Details JSON'];

function curationRegistrationKey_(value) {
  return String(value || '').replace(/\s+/g,' ').trim().toLowerCase();
}
function curationRegistrationHash_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(value))
    .map(function(b){return ('0'+((b+256)%256).toString(16)).slice(-2);}).join('');
}

// Pure coordinator; every side effect is supplied by a store. Intent precedes
// mutation. An ambiguous result never causes an automatic second append.
function curationRegistrationEngine_(store, action, input) {
  const context=store.context(input.eventName);
  const history=store.history(context.eventName);
  const active=store.active(context.eventName);
  const protectedLinks=(context.sourceSnapshot.existingCurationLinks||[]).concat(active);
  const latestById={};history.forEach(function(r){latestById[r.operationId]=r;});
  const operations=Object.keys(latestById).map(function(id){return latestById[id];});
  function result(state, extra) {return Object.assign({state:state,eventName:context.eventName,
    taskUrl:context.taskUrl,destination:'sandbox',sandboxUrl:store.url(),history:store.history(context.eventName)},extra||{});}
  function log(op, state, details) {
    const entry={operationId:op.operationId,eventName:context.eventName,action:action,state:state,
      actor:store.actor(),at:store.now(),details:Object.assign({operation:op},details||{})};
    if(JSON.stringify(entry.details).length>45000)throw new Error('Audit record exceeds supported size; nothing further was written.');
    store.audit(entry);return entry;
  }
  function original(id) {
    const intent=history.find(function(r){return r.operationId===id && r.state==='INTENT';});
    if(!intent)throw new Error('No Helper-owned registration was found.');
    return intent.details.operation;
  }
  if(action==='history') {
    const latest=history[history.length-1];
    const unresolved=operations.find(function(r){return /INTENT|UNCERTAIN|SOURCE_WRITTEN/.test(r.state);});
    const selected=unresolved||latest;
    return result(unresolved?'RECONCILIATION_REQUIRED':latest?latest.state:'HISTORY',selected?{operationId:selected.operationId,taskUrl:selected.details.operation.taskUrl}:{});
  }
  if(action==='preview') return result(protectedLinks.length?'EXISTING_LINK_PRESERVED':'PREVIEW',{
    canRegister:!protectedLinks.length && !operations.some(function(r){return /INTENT|UNCERTAIN|SOURCE_WRITTEN/.test(r.state);}),
    existingLinks:protectedLinks,previewKey:store.fingerprint({context:context,active:active}),
    message:'Test workspace only. Production inventory will not be changed.'});
  if(action==='register') {
    const unresolved=operations.find(function(r){return /INTENT|UNCERTAIN|SOURCE_WRITTEN/.test(r.state);});
    if(unresolved)return result('RECONCILIATION_REQUIRED',{operationId:unresolved.operationId});
    if(protectedLinks.length) return result('EXISTING_LINK_PRESERVED',{existingLinks:protectedLinks,
      operationId:active.length===1?active[0].operationId||'':''});
    if(input.previewKey!==store.fingerprint({context:context,active:active}))throw new Error('The preview changed. Preview again before recording.');
    const op={operationId:store.uuid(),eventName:context.eventName,taskUrl:context.taskUrl,
      actor:store.actor(),plannedAt:store.now(),sourceEventId:context.sourceEventId,
      prioritySetId:context.prioritySetId,sourceSnapshot:context.sourceSnapshot,
      before:active,destinationId:store.id()};
    log(op,'INTENT');
    try {
      const source=store.append(op);
      log(op,'SOURCE_WRITTEN',{source:source});
      if(!store.verify(op,false))throw new Error('Source or effective curation link did not verify.');
      log(op,'VERIFIED',{source:source});
      return result('VERIFIED',{operationId:op.operationId});
    } catch(error) {
      try {log(op,'UNCERTAIN',{error:String(error.message||error)});}catch(ignore){}
      return result('RECONCILIATION_REQUIRED',{operationId:op.operationId,error:String(error.message||error)});
    }
  }
  const op=original(input.operationId);
  if(op.destinationId!==store.id())throw new Error('This action belongs to another destination.');
  const latest=latestById[op.operationId];
  if(action==='withdraw') {
    if(latest.state==='WITHDRAWN')return result('WITHDRAWN',{operationId:op.operationId});
    if(latest.state!=='VERIFIED')return result('RECONCILIATION_REQUIRED',{operationId:op.operationId});
    const other=active.filter(function(r){return r.operationId!==op.operationId;});
    if(other.length || !store.verify(op,false)) {
      log(op,'WITHDRAW_REFUSED',{reason:'The owned source or another registration has changed.'});
      return result('CHANGED_RECORD_PRESERVED',{operationId:op.operationId});
    }
    log(op,'WITHDRAW_INTENT');
    try {
      const source=store.withdraw(op);
      log(op,'WITHDRAW_SOURCE_WRITTEN',{source:source});
      if(!store.verify(op,true))throw new Error('Withdrawal did not verify in source and curation view.');
      log(op,'WITHDRAWN',{source:source});return result('WITHDRAWN',{operationId:op.operationId});
    } catch(error) {
      try {log(op,'UNCERTAIN_WITHDRAW',{error:String(error.message||error)});}catch(ignore){}
      return result('RECONCILIATION_REQUIRED',{operationId:op.operationId,error:String(error.message||error)});
    }
  }
  if(action==='reconcile') {
    if(store.verify(op,true)) {log(op,'WITHDRAWN',{reconciled:true});return result('WITHDRAWN',{operationId:op.operationId});}
    if(store.verify(op,false)) {log(op,'VERIFIED',{reconciled:true});return result('VERIFIED',{operationId:op.operationId});}
    return result('RECONCILIATION_REQUIRED',{operationId:op.operationId,message:'No unambiguous verified outcome. No new source write was attempted.'});
  }
  throw new Error('Unsupported curation registration action.');
}

function curationRegistrationSourceContext_(eventName) {
  const report=getIntakeReport(eventName),receipt=checkWeaverEvent(eventName);
  if(!report.event || !receipt.curationTaskUrl || !receipt.transferComplete)throw new Error('A current, verified Weaver event connection is required.');
  const ss=getRetitleSpreadsheet_(),sheet=ss.getSheetByName('Form Responses 1'),last=sheet.getLastRow();
  const actions=last>1?sheet.getRange(2,2,last-1,1).getDisplayValues():[];
  const links=last>1?sheet.getRange(2,50,last-1,2).getDisplayValues():[];
  const existing=[];
  links.forEach(function(row,i){if(actions[i][0]==='Add Curation Form' && curationRegistrationKey_(row[0])===curationRegistrationKey_(report.eventName))existing.push({eventName:row[0],url:row[1],sourceRow:i+2});});
  return {eventName:report.eventName,taskUrl:receipt.curationTaskUrl,sourceEventId:receipt.eventId,prioritySetId:receipt.prioritySetId,
    sourceSnapshot:{spreadsheetId:RETITLE_CONFIG.spreadsheetId,eventRow:report.event.rowNumber,closed:report.closed,
      titles:report.titles.map(function(t){return {row:t.rowNumber,title:t.title,contentLink:t.contentLink};}),existingCurationLinks:existing}};
}

function setupCurationRegistrationSandbox(eventName) {
  weaverRequireOwner_();
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    const props=PropertiesService.getScriptProperties();
    if(props.getProperty(CURATION_REGISTRATION_SANDBOX_PROP))return curationRegistrationEngine_(curationRegistrationStore_(eventName,'history'),'history',{eventName:eventName});
    const context=curationRegistrationSourceContext_(eventName);
    const book=SpreadsheetApp.create('Footage Helper — curation registration sandbox');
    const source=book.getSheets()[0];source.setName('Form Responses 1');
    const headers=new Array(51).fill('');headers[0]='Timestamp';headers[1]='Select an option';headers[47]='Email Address';headers[49]='Enter Event (copy and paste so that it matches exactly)';headers[50]='Paste link to curation form';
    if(source.getMaxColumns()<51)source.insertColumnsAfter(source.getMaxColumns(),51-source.getMaxColumns());
    source.getRange(1,1,1,51).setValues([headers]);source.setFrozenRows(1);
    const view=book.insertSheet('Curation');view.getRange('A1:B1').setValues([['Event','Curation link']]);
    view.getRange('A2').setFormula('=IFERROR(FILTER({\'Form Responses 1\'!AX2:AX,\'Form Responses 1\'!AY2:AY},\'Form Responses 1\'!B2:B="Add Curation Form"),"")');
    const audit=book.insertSheet('Audit');audit.appendRow(CURATION_REGISTRATION_AUDIT_HEADERS);audit.setFrozenRows(1);
    // Seed existing registrations so the sandbox cannot pretend a legacy link is absent.
    context.sourceSnapshot.existingCurationLinks.forEach(function(link){const row=new Array(51).fill('');row[0]=new Date();row[1]='Add Curation Form';row[47]='sandbox baseline';row[49]=link.eventName;row[50]=link.url;source.appendRow(row);});
    props.setProperty(CURATION_REGISTRATION_SANDBOX_PROP,book.getId());
    return {state:'SANDBOX_CREATED',destination:'sandbox',sandboxUrl:book.getUrl(),eventName:context.eventName,history:[]};
  } finally {lock.releaseLock();}
}

function curationRegistrationStore_(eventName,action) {
  const id=PropertiesService.getScriptProperties().getProperty(CURATION_REGISTRATION_SANDBOX_PROP);
  if(!id)throw new Error('Create the test workspace first.');
  if(id===RETITLE_CONFIG.spreadsheetId)throw new Error('Production inventory is not an allowed destination.');
  const book=SpreadsheetApp.openById(id),source=book.getSheetByName('Form Responses 1'),audit=book.getSheetByName('Audit'),view=book.getSheetByName('Curation');
  if(!source||!audit||!view)throw new Error('Sandbox structure is incomplete.');
  const context=(action==='preview'||action==='register')?curationRegistrationSourceContext_(eventName):{eventName:eventName,taskUrl:'',sourceSnapshot:{existingCurationLinks:[]}};
  function rows(){const n=source.getLastRow();return n<2?[]:source.getRange(2,1,n-1,51).getValues().map(function(v,i){return {row:i+2,values:v,note:source.getRange(i+2,1).getNote()};});}
  function expected(op,withdrawn){const row=new Array(51).fill('');row[0]=new Date(op.plannedAt);row[1]=withdrawn?'Withdrawn Weaver curation link':'Add Curation Form';row[47]=op.actor;row[49]=op.eventName;row[50]=op.taskUrl;return row;}
  function own(op){const found=rows().filter(function(r){return r.note==='helper-curation:'+op.operationId || (r.values[0] instanceof Date && r.values[0].toISOString()===op.plannedAt && r.values[47]===op.actor && r.values[49]===op.eventName && r.values[50]===op.taskUrl);});if(found.length!==1)return null;return found[0];}
  function same(a,b){return JSON.stringify(a)===JSON.stringify(b);}
  return {id:function(){return id;},url:function(){return book.getUrl();},context:function(){return context;},
    actor:function(){return Session.getActiveUser().getEmail();},uuid:function(){return Utilities.getUuid();},
    now:function(){return new Date().toISOString();},fingerprint:curationRegistrationHash_,
    active:function(name){return rows().filter(function(r){return r.values[1]==='Add Curation Form'&&curationRegistrationKey_(r.values[49])===curationRegistrationKey_(name);}).map(function(r){return {row:r.row,eventName:r.values[49],url:r.values[50],operationId:r.note.replace(/^helper-curation:/,'')};});},
    history:function(name){const n=audit.getLastRow();return n<2?[]:audit.getRange(2,1,n-1,7).getValues().filter(function(v){return curationRegistrationKey_(v[1])===curationRegistrationKey_(name);}).map(function(v){return {operationId:v[0],eventName:v[1],action:v[2],state:v[3],actor:v[4],at:v[5],details:JSON.parse(v[6])};});},
    audit:function(entry){const values=[entry.operationId,entry.eventName,entry.action,entry.state,entry.actor,entry.at,JSON.stringify(entry.details)];audit.appendRow(values);SpreadsheetApp.flush();if(!same(audit.getRange(audit.getLastRow(),1,1,7).getValues()[0],values))throw new Error('Audit entry readback failed.');},
    append:function(op){source.appendRow(expected(op,false));SpreadsheetApp.flush();const r=own(op);if(!r)throw new Error('Written source row could not be uniquely identified.');source.getRange(r.row,1).setNote('helper-curation:'+op.operationId);return {row:r.row,after:r.values};},
    verify:function(op,withdrawn){const r=own(op);if(!r||!same(r.values,expected(op,withdrawn)))return false;SpreadsheetApp.flush();const n=view.getLastRow(),links=n<2?[]:view.getRange(2,1,n-1,2).getDisplayValues();const present=links.some(function(v){return v[0]===op.eventName&&v[1]===op.taskUrl;});return withdrawn?!present:present;},
    withdraw:function(op){const r=own(op);if(!r||!same(r.values,expected(op,false)))throw new Error('The owned source row changed; nothing was withdrawn.');source.getRange(r.row,2).setValue('Withdrawn Weaver curation link');SpreadsheetApp.flush();return {row:r.row,before:r.values,after:expected(op,true)};}
  };
}

function curationRegistrationRun_(action,eventName,extra) {
  weaverRequireOwner_();const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {return curationRegistrationEngine_(curationRegistrationStore_(eventName,action),action,Object.assign({eventName:eventName},extra||{}));}
  finally {lock.releaseLock();}
}
function previewCurationRegistration(eventName){return curationRegistrationRun_('preview',eventName);}
function registerCurationLinkSandbox(eventName,previewKey){return curationRegistrationRun_('register',eventName,{previewKey:previewKey});}
function withdrawCurationLinkSandbox(eventName,operationId){return curationRegistrationRun_('withdraw',eventName,{operationId:operationId});}
function reconcileCurationRegistration(eventName,operationId){return curationRegistrationRun_('reconcile',eventName,{operationId:operationId});}
function getCurationRegistrationHistory(eventName){return curationRegistrationRun_('history',eventName);}
