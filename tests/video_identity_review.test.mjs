import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {runInNewContext} from 'node:vm';
import {prepareVideoIdentityReview, assertVideoImportPreflight, buildWeaverVideoImport} from '../video_curation_handoff.mjs';
const clean=x=>String(x??'').trim();
const candidate={prioritySetId:'event',sourceFileId:'final',sourceFileName:'Poet - Title.mov',
  sourceMediaType:'FV',sourceReviewFileId:'',baseScore:8.05,excerptBonus:0,candidateScore:8.05,
  knownAssets:[{fileId:'final'},{fileId:'original'},{fileId:'replacement'}],
  reviewEvidence:[{sourceFileId:'original'},{sourceFileId:'replacement'}],
  gate:{prioritySetId:'event',sourceFileId:'final',videoRecordId:'weaver:video:final',
    decision:'ready_for_poetry_please',publishableAssetUrl:'https://drive.google.com/file/d/final/view',
    sourceEvent:'event',sourceEventLabel:'Event',selectedExcerptRecordIds:['excerpt-1'],
    baseScore:8.05,excerptBonus:1,candidateScore:9.05,decidedBy:'original@example.com',decidedAt:'original-time'}};
const payload={prioritySetId:'event',sourceFileId:'final',sourceReviewFileId:'original',
  evidenceUrl:'https://docs.google.com/document/d/evidence/edit',note:'Compared the reviewed recording and final version at three positions; the performance and cut match.'};
test('review requires a real stored review source, documentary evidence and an unsent approved gate',()=>{
  const review=prepareVideoIdentityReview(candidate,payload,'operator@example.com');
  assert.equal(review.sourceReviewFileId,'original');
  for(const c of [{...candidate,publicationRestricted:true},{...candidate,existingCanonicalVideoId:'existing'},
    {...candidate,gate:{...candidate.gate,decision:'hold'}},
    {...candidate,gate:{...candidate.gate,poetryPleaseHandoff:{status:'sending'}}}]){
    assert.throws(()=>prepareVideoIdentityReview(c,payload,'operator@example.com'));
  }
  for(const p of [{...payload,sourceReviewFileId:'unreviewed'},{...payload,note:'too short'},
    {...payload,evidenceUrl:'https://untrusted.example/evidence'},
    {...payload,evidenceUrl:'https://docs.google.com/anything'}]){
    assert.throws(()=>prepareVideoIdentityReview(candidate,p,'operator@example.com'));
  }
});
test('an already verified selection is idempotent and handoff scores cannot restore a legacy bonus',()=>{
  const gate={...candidate.gate,sourceReviewFileId:'original',identityReview:{status:'verified'},publishableAssetMediaType:'FV'};
  assert.equal(prepareVideoIdentityReview({...candidate,gate},payload,'operator@example.com').alreadyVerified,true);
  const record=buildWeaverVideoImport({...candidate,sourceReviewFileId:'original'},gate);
  assert.equal(record.sourceRecordId,'weaver:video:original');
  assert.equal(record.finalAssetUrl,candidate.gate.publishableAssetUrl);
  assert.equal(record.candidateScore,8.05);assert.equal(record.excerptBonus,0);
  assert.equal(gate.candidateScore,9.05);assert.equal(gate.decidedAt,'original-time');
});
test('only the exact old raw-source rejection exposes identity review; receiver attempts stay held',()=>{
  const prior={status:'blocked',error:'final_asset_matches_raw_source',ok:false,sourceRecordId:'weaver:video:final'};
  const c={...candidate,requiresManualIdentityReview:true,gate:{...candidate.gate,handoffStatus:'blocked',poetryPleaseHandoff:prior}};
  assert.equal(prepareVideoIdentityReview(c,payload,'operator').sourceReviewFileId,'original');
  const app=readFileSync(new URL('../public/app.js',import.meta.url),'utf8');
  const visible=runInNewContext(`(${app.match(/function canReviewVideoRecordingIdentity\([^]*?\n}\n/)[0]})`);
  assert.equal(visible(c),true);
  for(const change of [{attemptId:'started'},{canonicalVideoId:'existing'},{error:'other'},
    {sourceRecordId:'weaver:video:other'},{ok:true},{status:'sending'}]){
    const changed={...c,gate:{...c.gate,poetryPleaseHandoff:{...prior,...change}}};
    assert.throws(()=>prepareVideoIdentityReview(changed,payload,'operator'));
    assert.equal(visible(changed),false);
  }
});
const server=readFileSync(new URL('../server.mjs',import.meta.url),'utf8');
function harness({occupied='',driveMismatch=false}={}){
  const calls=[];
  const context={enqueueVideoPipeline:async()=>{calls.push({type:"queued"});},cleanSheetWhitespace:clean,prepareVideoIdentityReview,assertVideoImportPreflight,
    poetryPleaseApiUrl:'https://receiver.example/api',poetryPleaseApiKey:'test-key',AbortSignal,Date,
    getVideoCurationCandidates:async()=>({candidates:[structuredClone(candidate)]}),
    getDriveFileMetadata:async()=>({mimeType:'video/quicktime',name:driveMismatch?'Different.mov':candidate.sourceFileName}),
    fetch:async url=>{const sourceRecordId=new URL(url).searchParams.get('sourceRecordId');calls.push({type:'receipt',sourceRecordId});
      return {ok:true,json:async()=>({ok:true,sourceRecordId,received:sourceRecordId===occupied,
        canonicalVideoId:sourceRecordId===occupied?'existing':'',ledger:{count:0}})};},
    syncWeaverRuntimeDb:async(action,body)=>{calls.push({type:action,body});return {saved:true,record:{...candidate.gate,sourceReviewFileId:body.review.sourceReviewFileId}};},
    invalidateQueueSnapshots(){}
  };
  const resolve=runInNewContext(`(${server.match(/async function resolveVideoCurationIdentity\([^]*?\n}\n/)[0]})`,context);
  return {resolve,calls};
}
test('identity save checks every known receiver source and never imports a video',async()=>{
  const {resolve,calls}=harness();const r=await resolve(payload,'operator@example.com');
  assert.equal(r.ok,true);
  assert.deepEqual(calls.filter(c=>c.type==='receipt').map(c=>c.sourceRecordId).sort(),[
    'weaver:video:final','weaver:video:original','weaver:video:replacement']);
  const save=calls.find(c=>c.type==='resolve_video_curation_identity');
  assert.ok(save);assert.equal(save.body.expectedGate.decidedAt,'original-time');
  assert.equal(save.body.review.reviewedBy,'operator@example.com');
  assert.equal(calls.length,5);assert.equal(calls.filter(c=>c.type==="queued").length,1);
});
test('a receiver identity on any historical file or mismatching media stops the save',async()=>{
  for(const options of [{occupied:'weaver:video:replacement'},{occupied:'weaver:video:original'},{driveMismatch:true}]){
    const {resolve,calls}=harness(options);await assert.rejects(resolve(payload,'operator@example.com'));
    assert.equal(calls.some(c=>c.type==='resolve_video_curation_identity'),false);
  }
  assert.match(server,/"\/api\/admin\/video-curation-gates\/identity"/);
  const route=server.slice(server.indexOf('  if (url.pathname === "/api/admin/video-curation-gates/identity"'));
  assert.match(route.slice(0,650),/verifyAdministrativeCaller/);
});
