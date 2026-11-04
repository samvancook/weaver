// Fixed, user-requested repairs from Eric Tu's Oct 2–Nov 3 email thread.
// No caller-supplied cells/values. Owner-only; no row/media deletion or broad handler.
const ERIC_REPAIR_AUDIT_PROP='ERIC_REPAIRS_20261104_AUDIT';
const ERIC_REPAIR_EMAIL='https://mail.google.com/mail/u/0/#all/FMfcgzQhWfXrsZpvZRPkVDTggCrnSHpn';
const ERIC_DONTE_FOLDER='1vJZo1VaqPcCs-nEaLTF0HzjVfg5umzA1';
const ERIC_DONTE_FILES=[
 ['0B_-uqSs_rEYCdUxra1JFelBhWFE','0-bSNAdLucFo76vAw75L937w'],
 ['0B_-uqSs_rEYCYlloaGNydHJmUk0','0-35pZuke3_4emVpwTkNuFew'],
 ['0B_-uqSs_rEYCaW03ZHZaZU1nb00','0-I9c1r-PQ_yJaNj-XSk4Hyw'],
 ['0B_-uqSs_rEYCYW5ybnAwdUJjNkU','0-JBEbRmbTK2rhvJtSui9pLg'],
 ['0B_-uqSs_rEYCZG5iMlFDN21mSTA','0-NAZlG2iQ9koYzAq2HhcCJA'],
 ['0B_-uqSs_rEYCQ1EtY1lQNjBjMnc','0-EzPnrnGajTWROlH6MAc2Yw']
];
function ericAuditBook_(){
 const props=PropertiesService.getScriptProperties(),id=props.getProperty(ERIC_REPAIR_AUDIT_PROP);
 if(id){if(id===RETITLE_CONFIG.spreadsheetId)throw new Error('Audit must be separate.');return SpreadsheetApp.openById(id);}
 const book=SpreadsheetApp.create('Footage Helper — Eric reported repairs — 2026-11-04');const sh=book.getSheets()[0];sh.setName('Audit');sh.appendRow(['Operation ID','State','Time','Actor','Details JSON']);sh.setFrozenRows(1);props.setProperty(ERIC_REPAIR_AUDIT_PROP,book.getId());return book;
}
function ericLog_(book,id,state,details){
 const row=[id,state,new Date().toISOString(),Session.getActiveUser().getEmail(),JSON.stringify(details)];if(row[4].length>45000)throw new Error('Audit entry too large.');const sh=book.getSheetByName('Audit');sh.appendRow(row);SpreadsheetApp.flush();if(JSON.stringify(sh.getRange(sh.getLastRow(),1,1,5).getValues()[0])!==JSON.stringify(row))throw new Error('Audit readback failed.');
}
function ericHistory_(book){const sh=book.getSheetByName('Audit'),n=sh.getLastRow();return n<2?[]:sh.getRange(2,1,n-1,5).getValues().map(function(v){return {operationId:v[0],state:v[1],at:v[2],actor:v[3],details:JSON.parse(v[4])};});}
function ericSourceInvariant_(ss,patches){
 const source=ss.getSheetByName('Form Responses 1'),raw=source.getDataRange().getValues();
 patches.filter(function(p){return p.sheet==='Form Responses 1';}).forEach(function(p){raw[p.row-1][p.col-1]='[owned repair cell]';});
 const titles=ss.getSheetByName('All titles w/performer & Event'),filenames=titles.getRange(1,5,titles.getLastRow(),1).getValues();
 patches.filter(function(p){return p.generated;}).forEach(function(p){filenames[p.row-1][0]='[owned repair cell]';});
 return {source:curationRegistrationHash_(raw),filenames:curationRegistrationHash_(filenames),anchors:curationRegistrationHash_([titles.getRange(1,1,1,42).getFormulas(),ss.getSheetByName('Performer').getRange(1,1,2,12).getFormulas()])};
}
function ericPerformerNames_(ss){const sh=ss.getSheetByName('Performer');return sh.getRange(2,7,sh.getLastRow()-1,1).getDisplayValues().map(function(v){return v[0];}).filter(Boolean).sort();}
function ericPickerState_(form,ids){return (ids||['1479705159','1675652896']).slice(0,2).map(function(id){const values=curationRegistrationChoiceValues_(form,Number(id));return {id:id,hash:curationRegistrationHash_(values.slice().sort()),count:values.length,relevant:values.filter(function(n){return /Garner|Bauman|Tadesse|Symone/.test(n);})};});}
function ericRefreshPickers_(ss,form){
 const names=ericPerformerNames_(ss);
 ['1479705159','1675652896'].forEach(function(id){const item=form.getItemById(Number(id)),type=item.getType();if(type===FormApp.ItemType.MULTIPLE_CHOICE)item.asMultipleChoiceItem().setChoiceValues(names);else if(type===FormApp.ItemType.LIST)item.asListItem().setChoiceValues(names);else throw new Error('Performer picker type changed.');});
 const text=form.getItemById(1653265853).asTextItem();if(text.getTitle()!=='Performer List')throw new Error('Performer help item changed.');text.setHelpText(names.join('\n'));
 const after=ericPickerState_(form);if(after.some(function(p){return p.hash!==curationRegistrationHash_(names);})||text.getHelpText()!==names.join('\n'))throw new Error('Performer picker/help verification failed.');return after;
}
function ericPlan_(ss,form){
 const checkedColumns=ericValidateSourceSchema_(ss);
 const source=ss.getSheetByName('Form Responses 1'),data=source.getDataRange().getValues(),titles=ss.getSheetByName('All titles w/performer & Event'),patches=[];
 const donte=[];const sourceTitleRows=[];data.forEach(function(v,i){if(v[1]==='Add Title')sourceTitleRows.push(i+1);});
 function patch(sheet,a1,oldValue,newValue,caseId){const sh=ss.getSheetByName(sheet),cell=sh.getRange(a1),value=cell.getValue();if(cell.getFormula())throw new Error('Refusing a formula cell: '+a1);if(value!==oldValue&&value!==newValue)throw new Error('Unexpected current value: '+sheet+'!'+a1);patches.push({sheet:sheet,a1:a1,row:cell.getRow(),col:cell.getColumn(),before:value,after:newValue,caseId:caseId});}
 function titleRow(sourceRow){const r=sourceTitleRows.indexOf(sourceRow)+2;if(r<2||String(titles.getRange(r,2).getValue())!==String(data[sourceRow-1][12])||String(titles.getRange(r,4).getValue())!==String(data[sourceRow-1][46]))throw new Error('Derived title mapping changed at '+sourceRow);return r;}
 for(let i=0;i<6;i++){
  const row=6647+i,number=17+i,fileInfo=ERIC_DONTE_FILES[i],file=DriveApp.getFileByIdAndResourceKey(fileInfo[0],fileInfo[1]),parents=file.getParents();let inFolder=false;while(parents.hasNext()){if(parents.next().getId()===ERIC_DONTE_FOLDER)inFolder=true;}
  if(file.getName()!=='Clip #'+number+'.mov'||!/^video\//.test(file.getMimeType())||!inFolder||data[row-1][1]!=='Add Title'||data[row-1][46]!=='Donte Collins - Release @ Honey (Re)Release - September 2017')throw new Error('Donte file/event evidence changed for clip '+number);
  const url='https://drive.google.com/file/d/'+fileInfo[0]+'/view?usp=drivesdk&resourcekey='+fileInfo[1];
  patch('Form Responses 1','AW'+row,'https://drive.google.com/drive/folders/'+ERIC_DONTE_FOLDER,url,'donte-links');
  patch('Form Responses 1','O'+row,'',file.getName(),'donte-links');
  const dr=titleRow(row),oldName=titles.getRange(dr,5).getValue();if(oldName!==''&&oldName!=='Center Cam'&&oldName!==file.getName())throw new Error('Filename unexpectedly populated at '+dr);
  donte.push({sourceRow:row,titleRow:dr,url:url,name:file.getName(),filenameBefore:oldName});
 }
 if(data[29][1]!=='Add a Performer'||data[6973][1]!=='Add a Performer'||data[6821][1]!=='Add a Performer')throw new Error('Performer source action changed.');
 patch('Form Responses 1','C30','Krystal Garner','','garner-duplicate');patch('Form Responses 1','CH6974','Joe Bauman','','joe-duplicate');
 patch('Form Responses 1','CH6822','Miriam Tadesse','Mariam Tadesse','mariam-spelling');
 const mariam=[];[6823,6837,6839,6842].forEach(function(row){if(data[row-1][1]!=='Add Title')throw new Error('Mariam title action changed.');const dr=titleRow(row),v=titles.getRange(dr,1,1,42).getDisplayValues()[0];if(v[8]||v[34]||v[35]||v[38])throw new Error('Mariam has an FV/score dependency requiring a migration plan.');mariam.push({sourceRow:row,titleRow:dr,title:v[1],link:v[5],oldContentId:v[23]});patch('Form Responses 1','L'+row,'Miriam Tadesse','Mariam Tadesse','mariam-spelling');});
 const right=source.getRange('AT6871').getDisplayValue();if(right!=='All Button Authors, All features, All workshop instructors - contract is not stored with video releases')throw new Error('Known rights category changed.');
 if(data[6704][11]!=='Kyle Tran Myhre'||String(data[6704][48]).indexOf('0B_-uqSs_rEYCMFlQT1FpVnJDcDA')<0)throw new Error('Kyle identity changed.');
 patch('Form Responses 1','AT6705','No video release',right,'kyle-rights');
 const names=ericPerformerNames_(ss);if(names.indexOf('Symone Jenkins-Hawkins')>=0||names.filter(function(n){return n==='Symone Johnson-Hawkins';}).length!==1)throw new Error('Symone is no longer already resolved.');
 return {checkedColumns:checkedColumns,patches:patches,donte:donte,mariam:mariam,kyleTitleRow:titleRow(6705),beforeInvariant:ericSourceInvariant_(ss,patches.concat(donte.map(function(t){return {generated:true,row:t.titleRow};}))),sourceEmail:ERIC_REPAIR_EMAIL,pickerBefore:ericPickerState_(form),relevantNames:names.filter(function(n){return /Garner|Bauman|Tadesse|Symone/.test(n);})};
}
function ericVerify_(ss,form,plan,formIds){
 const errors=[];plan.patches.forEach(function(p){if(ss.getSheetByName(p.sheet).getRange(p.a1).getValue()!==p.after)errors.push(p.sheet+'!'+p.a1);});
 const unchanged=JSON.stringify(ericSourceInvariant_(ss,ericInvariantPatches_(plan)))===JSON.stringify(plan.beforeInvariant);if(!unchanged)errors.push('Unrelated source/filename/formula change');
 const names=ericPerformerNames_(ss),wanted=['Garner','Joel Bauman','Mariam Tadesse','Symone Johnson-Hawkins'],bad=['Krystal Garner','Joe Bauman','Miriam Tadesse','Symone Jenkins-Hawkins'];
 wanted.forEach(function(n){if(names.filter(function(v){return v===n;}).length!==1)errors.push('Canonical name count: '+n);});bad.forEach(function(n){if(names.indexOf(n)>=0)errors.push('Duplicate remains: '+n);});
 const titleSheet=ss.getSheetByName('All titles w/performer & Event');plan.donte.forEach(function(t){if(titleSheet.getRange(t.titleRow,6).getValue()!==t.url||titleSheet.getRange(t.titleRow,5).getValue()!==ss.getSheetByName('Form Responses 1').getRange(t.sourceRow,15).getValue())errors.push('Donte derived link: '+t.titleRow);});const rights=plan.patches.find(function(p){return p.caseId==='kyle-rights';});if(titleSheet.getRange(plan.kyleTitleRow,3).getValue()!==rights.after)errors.push('Kyle derived rights');plan.mariam.forEach(function(t){const row=titleSheet.getRange(t.titleRow,1,1,42).getDisplayValues()[0];if(row[0]!=='Mariam Tadesse'||row[1]!==t.title||row[5]!==t.link||row[8]||row[34]||row[35]||row[38])errors.push('Mariam derived dependency: '+t.titleRow);});
 const pickers=ericPickerState_(form,formIds);if(pickers.some(function(p){return p.hash!==curationRegistrationHash_(names);}))errors.push('Picker parity');
 return {verified:!errors.length,errors:errors,unrelatedSourceAndFormulasUnchanged:unchanged,relevantNames:names.filter(function(n){return /Garner|Bauman|Tadesse|Symone/.test(n);}),pickers:pickers,changedCells:plan.patches.filter(function(p){return p.before!==p.after;}).length};
}
function previewEricReportedRepairs(){weaverRequireOwner_();const ss=getRetitleSpreadsheet_(),form=FormApp.openById(CURATION_INVENTORY_FORM_ID),plan=ericPlan_(ss,form),book=ericAuditBook_();return {state:'PREVIEW',previewKey:curationRegistrationHash_(plan),auditUrl:book.getUrl(),plan:plan};}
function applyEricReportedRepairs(previewKey){
 weaverRequireOwner_();const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{const ss=getRetitleSpreadsheet_(),form=FormApp.openById(CURATION_INVENTORY_FORM_ID),book=ericAuditBook_(),history=ericHistory_(book);
  if(history.length)return {state:history[history.length-1].state,operationId:history[history.length-1].operationId,auditUrl:book.getUrl(),message:'An execution is already recorded. No repeat write attempted.'};
  const plan=ericPlan_(ss,form);if(curationRegistrationHash_(plan)!==previewKey)throw new Error('The repair preview changed.');return ericExecutePlan_(ss,form,book,plan);
 }finally{lock.releaseLock();}
}
function getEricReportedRepairReport(){weaverRequireOwner_();const book=ericAuditBook_(),history=ericHistory_(book);return {auditUrl:book.getUrl(),sourceEmail:ERIC_REPAIR_EMAIL,history:history,remaining:[{issue:'Donte duplicate clip 17 / missing clip 15',status:'Needs review',detail:'Earlier source row 6458 already references clip 17; do not guess its replacement.'},{issue:'Mariam folder link and repeated B**** title',status:'Needs review',detail:'Source row 6823 links a folder; row 6842 has a file with the same title. Verify whether this is a duplicate before merging.'},{issue:'Historical Krystal Garner title labels',status:'Alias migration deferred',detail:'This repair does not migrate the ten historical title labels/content IDs.'}]};}
function ericRepairPage_(){const r=getEricReportedRepairReport();function esc(s){return String(s).replace(/[&<>\"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}return '<style>body{font:16px system-ui;max-width:900px;margin:32px auto;padding:0 20px;line-height:1.5}h2{margin-top:32px}</style><h1>Eric’s reported footage fixes</h1><p>Owner repair report. No repair runs by opening this page.</p><p><a href="'+r.auditUrl+'">Open full before/after audit</a></p><p><a href="'+r.sourceEmail+'">Original email evidence</a></p>'+r.history.map(function(h){return '<p>'+esc(h.at)+' · '+esc(h.state)+' · '+esc(h.operationId)+'</p>';}).join('')+(r.history.length&&r.history[r.history.length-1].state==='VERIFIED'?'<h2>Verified repairs</h2><ul><li>Six Donte file links and source filenames corrected; generated filenames verified.</li><li>Krystal Garner and Joe Bauman duplicate selectable entries removed; history retained.</li><li>Mariam Tadesse corrected in the registry and four title records.</li><li>Kyle’s reported author-contract category recorded.</li><li>Symone’s prior correction rechecked.</li></ul>':'')+(!r.history.length?'<h2>Staged — no production changes applied</h2><ul><li>Correct six Donte file links and filename inputs in Form Responses 1; refresh their generated filename outputs.</li><li>Remove Krystal Garner and Joe Bauman duplicate selectable entries.</li><li>Correct Mariam Tadesse in the registry and four title records.</li><li>Record Kyle’s reported author-contract category.</li><li>Symone’s prior correction is already verified.</li></ul><p>Source-input plan: Form Responses 1 only. Generated filenames are populated from column O.</p>':'')+'<h2>Remaining review</h2>'+r.remaining.map(function(v){return '<h3>'+esc(v.issue)+'</h3><p>'+esc(v.status)+': '+esc(v.detail)+'</p>';}).join('');}

function recheckEricReportedRepairs(){
 weaverRequireOwner_();const book=ericAuditBook_(),history=ericHistory_(book),intent=history.find(function(h){return h.state==='INTENT';});if(!intent)throw new Error('No recorded repair intent.');const result=ericVerify_(getRetitleSpreadsheet_(),FormApp.openById(CURATION_INVENTORY_FORM_ID),intent.details.plan);if(result.verified)ericLog_(book,intent.operationId,'VERIFIED',{result:result,rechecked:true});return {state:result.verified?'VERIFIED':'UNCERTAIN',operationId:intent.operationId,auditUrl:book.getUrl(),result:result};
}
function undoEricReportedRepairs(operationId){
 weaverRequireOwner_();const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{return ericUndoPlan_(getRetitleSpreadsheet_(),FormApp.openById(CURATION_INVENTORY_FORM_ID),ericAuditBook_(),operationId);}finally{lock.releaseLock();}
}
