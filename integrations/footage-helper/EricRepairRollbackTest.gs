// A small isolated fixture; no production workbook or form is copied or written.
// It uses only the affected case inputs and the same apply/populate/undo code.
function testEricSourceRepairUndo(){
 weaverRequireOwner_();const lock=LockService.getScriptLock();lock.waitLock(30000);
 try{
  const prod=getRetitleSpreadsheet_(),prodForm=FormApp.openById(CURATION_INVENTORY_FORM_ID),plan=JSON.parse(JSON.stringify(ericPlan_(prod,prodForm)));
  const book=SpreadsheetApp.create('Footage Helper — source repair rollback proof'),source=book.getSheets()[0];source.setName('Form Responses 1');source.insertColumnsAfter(26,73);source.insertRowsAfter(1000,6002);
  const titles=book.insertSheet('All titles w/performer & Event');titles.insertColumnsAfter(26,16);titles.insertRowsAfter(1000,4281);
  const performer=book.insertSheet('Performer');book.insertSheet('Audit').appendRow(['Operation ID','State','Time','Actor','Details JSON']);
  const headers=prod.getSheetByName('Form Responses 1').getRange(1,1,1,99).getValues();source.getRange(1,1,1,99).setValues(headers);
  const rows=Array.from(new Set(plan.patches.map(function(p){return p.row;})));rows.forEach(function(row){const original=prod.getSheetByName('Form Responses 1').getRange(row,1,1,99).getValues()[0];source.getRange(row,1).setValue(new Date('2026-11-04T00:00:00Z'));[2,12,13,46,47,49].forEach(function(col){source.getRange(row,col).setValue(original[col-1]);});});
  plan.patches.forEach(function(p){source.getRange(p.a1).setValue(p.before);});
  [[40,'Garner'],[41,'Joel Bauman'],[42,'Symone Johnson-Hawkins']].forEach(function(v){source.getRange(v[0],2).setValue('Add a Performer');source.getRange(v[0],86).setValue(v[1]);});
  // Same registry formula as the live source, pointing only at this fixture.
  performer.getRange('A2').setFormula(prod.getSheetByName('Performer').getRange('A2').getFormula());performer.getRange('G2').setFormula('=UNIQUE(A2:A)');
  const all=plan.donte.concat(plan.mariam).concat([{sourceRow:6705,titleRow:plan.kyleTitleRow}]);
  all.forEach(function(t){const raw=source.getRange(t.sourceRow,1,1,99).getValues()[0],r=t.titleRow;titles.getRange(r,1).setFormula("='Form Responses 1'!L"+t.sourceRow);titles.getRange(r,2).setValue(raw[12]);titles.getRange(r,3).setFormula("='Form Responses 1'!AT"+t.sourceRow);titles.getRange(r,4).setValue(raw[46]);titles.getRange(r,6).setFormula("='Form Responses 1'!AW"+t.sourceRow);titles.getRange(r,24).setFormula('=A'+r+'&" - "&B'+r);});
  plan.donte.forEach(function(t,i){t.filenameBefore='Previous generated filename '+i;titles.getRange(t.titleRow,5).setValue(t.filenameBefore);});
  const form=FormApp.create('Footage Helper — rollback proof only');form.setAcceptingResponses(false);
  const a=form.addMultipleChoiceItem().setTitle('Original picker A').setChoiceValues(['Previous A','Previous B']).showOtherOption(true),b=form.addListItem().setTitle('Original picker B').setChoiceValues(['Old list entry']),help=form.addTextItem().setTitle('Performer List').setHelpText('Exact original help text');const ids=[a.getId(),b.getId(),help.getId()];
  SpreadsheetApp.flush();plan.beforeInvariant=ericSourceInvariant_(book,ericInvariantPatches_(plan));
  const before={source:curationRegistrationHash_(source.getDataRange().getValues()),titles:curationRegistrationHash_(titles.getDataRange().getValues()),formulas:curationRegistrationHash_([source.getDataRange().getFormulas(),titles.getDataRange().getFormulas(),performer.getDataRange().getFormulas()]),form:curationRegistrationHash_(ericFormSnapshot_(form,ids))};
  const applied=ericExecutePlan_(book,form,book,plan,{fixture:true,formIds:ids,verify:function(){return ericVerify_(book,form,plan,ids);}});if(applied.state!=='VERIFIED')throw new Error('Fixture apply failed: '+JSON.stringify(applied));
  const restored=ericUndoPlan_(book,form,book,applied.operationId);
  const after={source:curationRegistrationHash_(source.getDataRange().getValues()),titles:curationRegistrationHash_(titles.getDataRange().getValues()),formulas:curationRegistrationHash_([source.getDataRange().getFormulas(),titles.getDataRange().getFormulas(),performer.getDataRange().getFormulas()]),form:curationRegistrationHash_(ericFormSnapshot_(form,ids))};
  const exact=JSON.stringify(before)===JSON.stringify(after);if(!exact)throw new Error('Fixture did not restore exactly: '+JSON.stringify({before:before,after:after}));
  const result={state:'ROLLBACK_PROVED',spreadsheetUrl:book.getUrl(),formUrl:form.getEditUrl(),operationId:applied.operationId,sourceInputCells:plan.patches.length,generatedFilenameOutputs:plan.donte.length,sourceRestored:before.source===after.source,titlesRestored:before.titles===after.titles,allFormulasRestored:before.formulas===after.formulas,exactFormSettingsRestored:before.form===after.form,formAcceptingResponses:form.isAcceptingResponses(),productionWritten:false};
  PropertiesService.getScriptProperties().setProperty('ERIC_SOURCE_ROLLBACK_PROOF',JSON.stringify(result));return result;
 }finally{lock.releaseLock();}
}
