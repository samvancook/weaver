import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const server = fs.readFileSync(new URL('../server.mjs', import.meta.url), 'utf8');
const app = fs.readFileSync(new URL('../public/app.js', import.meta.url), 'utf8');
function fn(source, name) {
  const match = source.match(new RegExp(`^(?:async )?function ${name}\\([\\s\\S]*?^}\\n`, 'm'));
  assert.ok(match, name); return match[0];
}
function row(id, author = 'Poet', excerpt = 'Original excerpt') {
  const cells = Array(65).fill('');
  Object.assign(cells, {24:id,25:author,26:'Poem',27:excerpt,28:'Book'});
  return cells;
}
function harness(rows = new Map([[2,row('source-2')],[3,row('source-3')]])) {
  const calls = [];
  const context = {
    Map,Set,Number,String,Object,Error,Promise, sourceSheetName:'Source',appVersion:'test',
    SHEET_SOURCE_CONFIG:vm.runInNewContext('('+server.match(/const SHEET_SOURCE_CONFIG = (\{[\s\S]*?^\});/m)[1]+')'),
    cleanSheetWhitespace: value=>String(value||'').trim().replace(/\s+/g,' '),
    fetchSheetValuesServer: async range=> { const n=Number(range.match(/!A(\d+)/)[1]);calls.push(['read',n]);return rows.has(n)?[rows.get(n)]:[]; },
    getCanonicalGraphicsBookAuthorMap: async()=>new Map(),parseIntakeMetadataFromNotes:()=>({}),
    buildPendingRecordFromSheetRow: cells=>({recordId:cells[24],author:cells[25],title:cells[26],excerptText:cells[27],bookTitle:cells[28]}),
    buildCorrectionRecordFromSheetRow: ()=>null,resolveGraphicsAuthor:v=>v,
    resolvePublishingBookMeta:()=>({bookShortener:'BOOK',releaseCatalog:'Fall'}),normalizeExcerptContentType:v=>v,
    extractGoogleDriveFileId:()=>'',normalizeCurationRating:()=>'',normalizeOptionalCurationScore:()=>null,
    buildReviewWriteRanges:update=>[{range:`Source!BC${update.sourceRow}`,values:[[update.reviewDecision]]}],
    buildExcerptRuntimeRecord:r=>r,
    batchUpdateSheetValuesServer:async requests=>{
      calls.push(['write',requests]);
      for(const request of requests){ const [,letters,n]=request.range.match(/!([A-Z]+)(\d+)$/);const column=Array.from(letters).reduce((v,ch)=>v*26+ch.charCodeAt(0)-64,0);rows.get(Number(n))[column-1]=request.values[0][0]; }
    },
    shadowExcerptRecordsToRuntime:async records=>{calls.push(['ledger',records]);return {ok:true};},
    syncAcceptedExcerptHandoffs:async records=>{calls.push(['queue',records]);return {ok:true};}
  };
  vm.createContext(context);
  vm.runInContext(['toA1Column','buildReviewWriteRanges','isTruthyParam','isAcceptedExcerptReviewDecision','normalizeDecision','normalizeReviewDecisionValue','getSheetExcerptReviewDecision','verifySavedReviewUpdates','validateReviewBatchTargets','verifyReviewSourceIdentity','enrichAcceptedExcerptUpdate','saveReviewsToSheets','saveSingleReviewToSheets'].map(n=>fn(server,n)).join('\n'),context);
  return {context,calls,rows};
}
const update=(sourceRow=2,recordId=`source-${sourceRow}`)=>({sourceRow,recordId,reviewDecision:'accept_skip_graphic'});

test('a moved source rejects the whole batch before any write, even if client supplies complete metadata', async()=>{
  const h=harness();h.rows.set(3,row('different-source'));
  await assert.rejects(h.context.saveReviewsToSheets([update(),{...update(3),excerptText:'Client text',bookTitle:'Book',bookShortener:'BOOK',releaseCatalog:'Fall'}]),/different source/);
  assert.ok(h.calls.some(c=>c[0]==='read'&&c[1]===3));
  assert.ok(!h.calls.some(c=>['write','ledger','queue'].includes(c[0])));
});

test('missing rows, duplicate targets and malformed targets never write a partial batch',async()=>{
  for(const batch of [[update(),update(9)], [update(),update(2,'other')], [update(),update(3,'source-2')], [{...update(),sourceRow:'2junk'}], [{...update(),sourceRow:2.5}], [{...update(),reviewDecision:'unknown'}]]) {
    const h=harness();await assert.rejects(h.context.saveReviewsToSheets(batch));
    assert.ok(!h.calls.some(c=>['write','ledger','queue'].includes(c[0])));
  }
});

test('legacy row identities require the exact loaded content and remain usable without a new approval step',async()=>{
  const h=harness(new Map([[2,row('')]]));const source={author:'Poet',title:'Poem',bookTitle:'Book',excerptText:'Original excerpt'};
  await assert.rejects(h.context.saveSingleReviewToSheets(update(2,'weaver:row-2')),/fresh load/);
  await assert.rejects(h.context.saveSingleReviewToSheets({...update(2,'weaver:row-2'),expectedSource:{...source,author:'Different poet'}}),/changed/);
  assert.ok(!h.calls.some(c=>c[0]==='write'));
  const result=await h.context.saveSingleReviewToSheets({...update(2,'weaver:row-2'),expectedSource:source});
  assert.equal(result.ok,true);assert.equal(result.recordId,'weaver:row-2');
});

test('caption-only decisions clear both graphic aliases on the server before ledger and handoff',async()=>{
  const h=harness();const result=await h.context.saveReviewsToSheets([{...update(),author:'Untrusted author',excerptText:'Client text',useForQi:true,graphicsQi:'1',useForInt:true,photos:'1'}]);
  assert.equal(result.savedCount,1);
  const saved=h.calls.find(c=>c[0]==='ledger')[1][0];
  for(const key of ['useForQi','graphicsQi','useForInt','photos'])assert.equal(saved[key],false,key);
  assert.equal(saved.recordId,'source-2');assert.equal(saved.excerptText,'Original excerpt');
  assert.equal(saved.author,'Poet');
  assert.equal(h.calls.filter(c=>c[0]==='write').length,1);
});

test('review updates retain the current source intake lane and collector rather than client defaults',async()=>{
  const h=harness();h.rows.get(2)[1]='collector@buttonpoetry.com';h.rows.get(2)[2]='Add a quote from a video';
  await h.context.saveReviewsToSheets([{...update(),mode:'book',intakeMode:'book',email:'changed@elsewhere.test'}]);
  const saved=h.calls.find(c=>c[0]==='ledger')[1][0];
  assert.equal(saved.mode,'video');assert.equal(saved.intakeMode,'video');
  assert.equal(saved.email,'collector@buttonpoetry.com');assert.equal(saved.submittedBy,'collector@buttonpoetry.com');
});

test('a same-ID content change rejects the loaded decision rather than approving unseen text',async()=>{
  const h=harness();const expectedSource={author:'Poet',title:'Poem',bookTitle:'Book',excerptText:'Old displayed text'};
  await assert.rejects(h.context.saveReviewsToSheets([{...update(),expectedSource}]),/displayed review source changed/);
  assert.ok(!h.calls.some(c=>['write','ledger','queue'].includes(c[0])));
});

test('a source ledger failure is reported separately and never queues downstream delivery',async()=>{
  for(const single of [false,true]) {
    const h=harness();h.context.shadowExcerptRecordsToRuntime=async()=>({ok:false,error:'unavailable'});
    const result=single?await h.context.saveSingleReviewToSheets(update()):await h.context.saveReviewsToSheets([update()]);
    assert.equal(result.ok,true);assert.equal(result.excerptHandoffs.ok,false);
    assert.match(result.excerptHandoffs.error,/decision.*saved.*ledger.*not.*synchronize/i);
    assert.ok(h.calls.some(c=>c[0]==='write'));assert.ok(!h.calls.some(c=>c[0]==='queue'));
  }
});

test('batch responses must confirm the exact number before UI treats a save as complete',async()=>{
  const context={adminFetch:async()=>({ok:true,json:async()=>({ok:true,savedCount:1})})};
  vm.createContext(context);vm.runInContext(fn(app,'requestBatchSave')+'\n'+fn(app,'buildBatchReviewSavePayload'),context);
  await assert.rejects(context.requestBatchSave([update(),update(3)]),/confirm every requested record/);
  context.adminFetch=async()=>({ok:true,json:async()=>({ok:true,savedCount:2})});
  assert.equal((await context.requestBatchSave([update(),update(3)])).savedCount,2);
  const expectedSource={author:'Poet',title:'Poem',bookTitle:'Book',excerptText:'Original excerpt'};
  assert.equal(context.buildBatchReviewSavePayload({...update(),expectedSource}).expectedSource,expectedSource);
});


test('write readback rejects changed source identity or unpersisted values before ledger or delivery',async()=>{
  for(const mode of ['identity','values']) {
    const h=harness();h.context.batchUpdateSheetValuesServer=async()=>{
      h.calls.push(['write']);if(mode==='identity')h.rows.get(2)[24]='moved-source';
    };
    await assert.rejects(h.context.saveReviewsToSheets([update()]),/Some decisions may already be saved/);
    assert.ok(!h.calls.some(c=>['ledger','queue'].includes(c[0])));
  }
});


test('trailing blank source cells do not look like an identity change after review columns are saved',async()=>{
  const h=harness(new Map([[2,row('source-2').slice(0,28)]]));
  assert.equal((await h.context.saveSingleReviewToSheets(update())).ok,true);
});


test('new exclusions and conflicting decisions are preserved instead of overwritten by a stale card',async()=>{
  for(const mode of ['excluded','rejected']) {
    const h=harness();if(mode==='excluded')h.rows.get(2)[29]='Y';else h.rows.get(2)[54]='REJECT';
    const expectedSource={author:'Poet',title:'Poem',bookTitle:'Book',excerptText:'Original excerpt',reviewDecision:''};
    await assert.rejects(h.context.saveReviewsToSheets([{...update(),expectedSource}]),/excluded|Another decision/);
    assert.ok(!h.calls.some(c=>['write','ledger','queue'].includes(c[0])));
  }
});


test('correction return saves overrides, clears correction/approval state and never approves or queues the excerpt',async()=>{
  const h=harness();h.rows.get(2)[54]='NEEDS_CORRECTION';h.rows.get(2)[35]='NEEDS_CORRECTION';
  vm.runInContext(fn(server,'buildExcerptRuntimeRecord'),h.context);
  h.context.syncAcceptedExcerptHandoffs=async records=>{
    const handoffs=records.map(r=>h.context.buildAcceptedExcerptHandoff(r)).filter(Boolean);
    assert.equal(handoffs.length,0);return {ok:true};
  };
  vm.runInContext(fn(server,'buildAcceptedExcerptHandoff'),h.context);
  const result=await h.context.saveSingleReviewToSheets({...update(),reviewDecision:'return_to_review',correctedTitle:'Correct title',useForQi:true,useForInt:true});
  assert.equal(result.ok,true);assert.equal(h.rows.get(2)[54],'');assert.equal(h.rows.get(2)[35],'');
  assert.equal(h.rows.get(2)[34],'N');assert.equal(h.rows.get(2)[55],'');assert.equal(h.rows.get(2)[57],'Correct title');
  const ledger=h.calls.find(c=>c[0]==='ledger')[1][0];
  assert.equal(ledger.reviewDecision,'');assert.equal(ledger.needsCorrection,false);assert.equal(ledger.approvedForUse,false);assert.equal(ledger.approvedForGraphics,false);
  assert.equal(ledger.correctedPoemTitle,'Correct title');
});

test('return-to-review cannot clear another record decision and UI No decision outside corrections stays unchanged',async()=>{
  const h=harness();await assert.rejects(h.context.saveReviewsToSheets([{...update(),reviewDecision:'return_to_review'}]),/Only a current correction/);
  assert.ok(!h.calls.some(c=>c[0]==='write'));
  const context={elements:{correctionList:{}},collectUpdatesFromContainer:()=>[{sourceRow:2,reviewDecision:'',useForQi:true,useForInt:true},{sourceRow:3,reviewDecision:'reject'}]};
  vm.createContext(context);vm.runInContext(fn(app,'collectCorrectionUpdates'),context);
  const updates=context.collectCorrectionUpdates();assert.equal(updates[0].reviewDecision,'return_to_review');assert.equal(updates[0].useForQi,false);assert.equal(updates[0].useForInt,false);assert.equal(updates[1].reviewDecision,'reject');
});

test('returned correction overrides appear in pending and approved reads while original source cells remain intact',()=>{
  const h=harness();Object.assign(h.context,{
    isSheetYes:v=>v==='Y',isPendingSheetReview:v=>!v,buildCatalogValidationPayload:()=>null,
    parseChicagoSheetTimestampToIso:v=>v,parseSheetInteger:v=>Number(v)||0,
    normalizeExcerptHandoffStatus:v=>v,resolveExcerptPoemKey:()=>({poemKey:'key'}),
    getCanonicalGraphicsBookAuthorMap:()=>new Map(),buildBrowserCatalogValidationPayload:()=>null,
    countWordsFromText:s=>s.split(/\s+/).length
  });
  const cells=row('source-2');Object.assign(cells,{56:'Correct author',57:'Correct title',58:'Correct book',59:'Correct excerpt'});
  const original=JSON.stringify(cells.slice(24,29));
  vm.runInContext(fn(server,'buildPendingRecordFromSheetRow')+'\n'+fn(server,'buildApprovedExcerptExportRecordFromSheetRow'),h.context);
  const pending=h.context.buildPendingRecordFromSheetRow(cells,0);
  assert.equal(pending.title,'Correct title');assert.equal(pending.author,'Correct author');assert.equal(pending.excerptText,'Correct excerpt');assert.equal(pending.bookTitle,'Correct book');
  assert.equal(pending.rawExcerptText,'Original excerpt');assert.equal(pending.rawTitle,'Poem');
  cells[54]='ACCEPT_SKIP_GRAPHIC';
  const approved=h.context.buildApprovedExcerptExportRecordFromSheetRow(cells,0);
  assert.equal(approved.poemTitle,'Correct title');assert.equal(approved.author,'Correct author');assert.equal(approved.excerptText,'Correct excerpt');assert.equal(approved.bookTitle,'Correct book');
  assert.equal(JSON.stringify(cells.slice(24,29)),original);
  cells[54]='';
  vm.runInContext(fn(app,'buildPendingRecordFromSheetRow'),h.context);
  const browser=h.context.buildPendingRecordFromSheetRow(cells,0);
  for(const key of ['title','author','bookTitle','excerptText','rawTitle','rawExcerptText'])assert.equal(browser[key],pending[key],key);
});
