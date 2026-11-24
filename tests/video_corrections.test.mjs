import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
const server=fs.readFileSync(new URL('../server.mjs',import.meta.url),'utf8');
const app=fs.readFileSync(new URL('../public/app.js',import.meta.url),'utf8');
const config=vm.runInNewContext('('+server.match(/const SHEET_SOURCE_CONFIG = (\{[\s\S]*?^\});/m)[1]+')');
const browserConfig=vm.runInNewContext('('+app.match(/const SHEET_SOURCE_CONFIG = (\{[\s\S]*?^\});/m)[1]+')');
function fn(source,name){const m=source.match(new RegExp(`^(?:async )?function ${name}\\([\\s\\S]*?^}\\n`,'m'));assert.ok(m,name);return m[0];}
function row(id,{video=true,event='Event / A & friends',book='',excluded=false,decision='NEEDS_CORRECTION'}={}){
  const r=Array(65).fill('');r[2]=video?'Add a quote from a video':'Add a quote from a book';
  r[8]='https://drive.google.com/file/d/video-file/view';r[9]='Poet';r[11]='Shared title';r[12]='Original words. Sentence.';r[17]=event;
  for(const[k,v]of Object.entries({recordId:id,author:'Poet',title:'Shared title',excerpt:'Original words. Sentence.',bookTitle:book,
    correctedExcerpt:'Original words, sentence.',excerptReviewDecision:decision,exclude:excluded?'Y':'',poemKey:'KEY-1'}))r[config.columnMap[k]-1]=v;
  return r;
}
function harness(browser,rows){
  const ranges=[];const c={SHEET_SOURCE_CONFIG:browser?browserConfig:config,REVIEW_VIDEO_EVENT_KEY:'__video_event__:',
    sourceSheetName:'Source',appVersion:'test',runtimeConfig:{sourceSheetName:'Source',appVersion:'test'},
    cleanSheetWhitespace:v=>String(v||'').trim().replace(/\s+/g,' '),normalizeBookKey:v=>String(v||'').trim().toLowerCase(),
    choosePreferredBookTitle:(a,b)=>a||b,resolveGraphicsAuthor:v=>v,isSheetYes:v=>v==='Y',parseSheetInteger:v=>Number(v)||0,
    countWordsFromText:v=>v.trim().split(/\s+/).length,getSheetExcerptReviewDecision:r=>r[54],isPendingSheetReview:v=>!v,
    parseIntakeMetadataFromNotes:v=>({sourceVideoUrl:v}),extractGoogleDriveFileId:v=>v?.match(/\/d\/([^/]+)/)?.[1]||'',
    buildCatalogValidationPayload:()=>({}),buildBrowserCatalogValidationPayload:()=>({}),
    getSourceSheetValuesCached:async()=>rows,getCanonicalGraphicsBookAuthorMap:async()=>new Map(),
    fetchSheetValuesServer:async range=>{ranges.push(range);return rows;},fetchSheetValues:async range=>{ranges.push(range);return rows;}
  };
  vm.createContext(c);
  const names=browser?['buildExcerptPayloadFromSheetRow','summarizeSimpleBooks','summarizeCorrectionSelections','matchesCorrectionSelection','loadCorrectionBooksFromSheetsFallback','loadCorrectionsFromSheetsFallback']
    :['buildCorrectionRecordFromSheetRow','summarizeBooks','summarizeCorrectionSelections','matchesCorrectionSelection','getCorrectionBooksFromSheets','getCorrectionsForBookFromSheets'];
  vm.runInContext(names.map(n=>fn(browser?app:server,n)).join('\n'),c);
  return {c,ranges,books:()=>browser?c.loadCorrectionBooksFromSheetsFallback():c.getCorrectionBooksFromSheets(),
    records:key=>browser?c.loadCorrectionsFromSheetsFallback(key):c.getCorrectionsForBookFromSheets(key),
    build:(r,i=0)=>browser?c.buildExcerptPayloadFromSheetRow(r,i):c.buildCorrectionRecordFromSheetRow(r,i)};
}
for(const browser of [false,true]){
  const lane=browser?'Sheets fallback':'primary API';
  test(`${lane}: a video correction without a book retains its source identity and entered override`,()=>{
    const h=harness(browser,[]),record=h.build(row('source-one'));
    assert.ok(record);assert.equal(record.bookTitle,'');assert.equal(record.intakeMode,'video');
    assert.equal(record.sourceEvent,'Event / A & friends');assert.equal(record.sourceVideoFileId,'video-file');
    assert.equal(record.recordId,'source-one');assert.equal(record.poemKey,'KEY-1');
    assert.equal(record.rawExcerptText,'Original words. Sentence.');assert.equal(record.excerptText,'Original words, sentence.');
    assert.equal(record.correctedExcerpt,'Original words, sentence.');
    const legacy=row('legacy');legacy[config.columnMap.author-1]='';legacy[config.columnMap.title-1]='';legacy[config.columnMap.excerpt-1]='';
    assert.equal(h.build(legacy).rawAuthor,'Poet');assert.equal(h.build(legacy).rawTitle,'Shared title');
    assert.equal(h.build(legacy).rawExcerptText,'Original words. Sentence.');
  });
  test(`${lane}: event choices select only their exact video records and leave book correction selection intact`,async()=>{
    const h=harness(browser,[row('one'),row('two',{event:'Other event'}),row('book',{video:false,book:'A Book'})]);
    const books=(await h.books()).books;const group=books.find(x=>x.value==='__video_event__:'+encodeURIComponent('Event / A & friends'));
    assert.ok(group);assert.equal(group.count,1);assert.ok(books.some(x=>x.title==='A Book'&&x.count===1));
    const selected=(await h.records(group.value)).excerpts;
    assert.equal(selected.length,1);assert.equal(selected[0].recordId,'one');assert.equal(selected[0].bookTitle,'');
    assert.equal((await h.records('A Book')).excerpts[0].recordId,'book');
    assert.ok(h.ranges.every(range=>range.endsWith('A2:BM')),'Poem keys must be read through column BM');
  });
  test(`${lane}: excluded, approved and non-correction video sources are never offered as correction work`,async()=>{
    const h=harness(browser,[row('pending'),row('excluded',{excluded:true}),row('approved',{decision:'ACCEPT'}),row('undecided',{decision:''})]);
    const group=(await h.books()).books.find(x=>x.value);assert.equal(group.count,1);
    const records=(await h.records(group.value)).excerpts;assert.equal(records.length,1);assert.equal(records[0].recordId,'pending');
  });
}
test('correction selector preserves encoded event values instead of treating the display label as a book',async()=>{
  const selection='__video_event__:'+encodeURIComponent('Event A');const options=[];
  const select={value:selection,appendChild:o=>options.push(o),innerHTML:''};
  const c={elements:{correctionBookSelect:select,correctionBookCountBadge:{}},document:{createElement:()=>({})},setStatus:()=>{},
    requestReviewApi:async()=>({books:[{title:'Video excerpts — Event A',value:selection,count:1}]}),runtimeConfig:{}};
  vm.createContext(c);vm.runInContext(fn(app,'loadCorrectionBooks'),c);await c.loadCorrectionBooks();
  assert.equal(options[1].value,selection);assert.equal(select.value,selection);
});
test('loading a video correction preserves empty book metadata and uses a friendly event label',async()=>{
  const record={recordId:'one',intakeMode:'video',bookTitle:''};const calls=[];
  const c={elements:{correctionBookSelect:{value:'__video_event__:Event%20A',selectedOptions:[{textContent:'Video excerpts — Event A (1)'}]}},
    currentCorrectionExcerpts:[],setStatus:s=>calls.push(['status',s]),requestReviewApi:async()=>({excerpts:[record]}),
    loadCatalogValidation:async rows=>calls.push(['catalog',rows]),renderCorrectionExcerpts:rows=>calls.push(['render',rows])};
  vm.createContext(c);vm.runInContext(fn(app,'loadCorrections'),c);await c.loadCorrections();
  assert.equal(c.currentCorrectionExcerpts[0].bookTitle,'');assert.equal(calls.find(x=>x[0]==='catalog')[1].length,0);
  assert.match(calls.filter(x=>x[0]==='status').at(-1)[1],/Video excerpts — Event A/);
  assert.ok(!calls.some(x=>x[0]==='status'&&x[1].includes('__video_event__')));
});
function actorRoute(){
  const writes=[];const start=server.indexOf('  if (url.pathname === "/api/save-graphics-qc"');
  const route=server.slice(start,server.indexOf('\n  if (url.pathname',start+1));
  const c={cleanSheetWhitespace:v=>String(v||'').trim(),adminAuthError:code=>new Error(code),
    readRequestBody:async req=>JSON.stringify(req.payload),saveGraphicsQcToSheets:async updates=>{writes.push(updates);return {ok:true};},
    invalidateQueueSnapshots:()=>{},sendJson:(_res,status,body)=>({status,body})};
  vm.createContext(c);vm.runInContext(fn(server,'attributeGraphicsQcUpdates')+`\nasync function handle(req){const url={pathname:'/api/save-graphics-qc'};const res={};${route}}`,c);
  return {c,writes};
}
test('QC route records the authenticated principal and overrides an attempted client reviewer impersonation',async()=>{
  const h=actorRoute();const result=await h.c.handle({method:'POST',weaverAdmin:{email:'Reviewer@buttonpoetry.com'},
    payload:{updates:[{pigCompletionId:'one',qcDecision:'reject',reviewedBy:'forged@elsewhere.test'}]}});
  assert.equal(result.status,200);assert.equal(h.writes.length,1);assert.equal(h.writes[0][0].reviewedBy,'reviewer@buttonpoetry.com');
  assert.equal(h.writes[0][0].pigCompletionId,'one');
});
test('QC route cannot persist a client-supplied reviewer when the verified principal is absent',async()=>{
  const h=actorRoute();const result=await h.c.handle({method:'POST',payload:{updates:[{reviewedBy:'forged@elsewhere.test'}]}});
  assert.equal(h.writes.length,0);assert.equal(result.body.ok,false);assert.match(result.body.error,/authentication_required/);
});
