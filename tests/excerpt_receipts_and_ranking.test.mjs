import {videoFileMediaType} from '../video_media_classification.mjs';
import {videoReviewProgress} from '../video_curation_handoff.mjs';
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { runInNewContext } from 'node:vm';

const server = readFileSync(new URL('../server.mjs', import.meta.url), 'utf8');
const clean = value => String(value ?? '').trim();
const fn = name => server.match(new RegExp(`(?:async )?function ${name}\\([^]*?\\n}\\n`))[0];

async function statuses(response) {
  const writes = [];
  const save = runInNewContext(`(${fn('updateExcerptHandoffStatuses')})`, {
    cleanSheetWhitespace: clean,
    SHEET_SOURCE_CONFIG: { startRow: 2, columnMap: { excerptPoetryPleaseStatus: 1,
      excerptPoetryPleaseUpdatedAt: 2, excerptPoetryPleaseNote: 3 } },
    sourceSheetName: 'Source', toA1Column: n => String.fromCharCode(64+n),
    batchUpdateSheetValuesServer: async rows => writes.push(...rows), invalidateQueueSnapshots() {},
    syncWeaverRuntimeDb: async (_action, payload) => ({ ok: true, records: payload.handoffs })
  });
  const records = [
    { recordId: 'a', payload: { sourceRow: 2 }, poetryPleaseItemId: 'retained-item', handedOffAt: 'retained-time' },
    { recordId: 'b', payload: { sourceRow: 3 } }
  ];
  return { ...(await save(records, response)), writes };
}

test('failed or incomplete batch responses never mark unacknowledged excerpts sent', async () => {
  for (const response of [undefined, { ok: false, error: 'offline' },
    { ok: true, results: [] }, { ok: true, results: [{ recordId: 'a', ok: true }] },
    { results: [{ recordId: 'other', ok: true, poetryPleaseItemId: 'wrong' }] },
    { results: [{ recordId: 'a', ok: true, poetryPleaseItemId: 'x' },
      { recordId: 'a', ok: true, poetryPleaseItemId: 'y' }] }]) {
    const result = await statuses(response);
    assert.ok(result.records.every(r => r.handoffStatus === 'failed'));
    assert.equal(result.records[0].poetryPleaseItemId, 'retained-item');
    assert.equal(result.records[0].handedOffAt, 'retained-time');
    assert.ok(result.records.every(r => r.errorMessage));
    assert.match(result.writes.find(w => w.range.endsWith('C2')).values[0][0], /retained-item; /);
  }
});

test('a partial batch marks only the exact acknowledged item sent', async () => {
  const result = await statuses({ ok: false, results: [
    { recordId: 'a', ok: true, poetryPleaseItemId: 'canonical-a' },
    { recordId: 'b', ok: false, error: 'metadata missing' }
  ] });
  assert.equal(result.records[0].handoffStatus, 'sent');
  assert.equal(result.records[0].poetryPleaseItemId, 'canonical-a');
  assert.equal(result.records[1].handoffStatus, 'failed');
  assert.equal(result.records[1].errorMessage, 'metadata missing');
});

test('durable persistence precedes the compatibility sheet and a storage failure stops it', async () => {
  const calls=[];
  const context={cleanSheetWhitespace:clean,SHEET_SOURCE_CONFIG:{startRow:2,columnMap:{excerptPoetryPleaseStatus:1,
    excerptPoetryPleaseUpdatedAt:2,excerptPoetryPleaseNote:3}},sourceSheetName:'Source',toA1Column:n=>String(n),
    invalidateQueueSnapshots(){},batchUpdateSheetValuesServer:async()=>{calls.push('sheet');throw new Error('offline');},
    syncWeaverRuntimeDb:async(_action,payload)=>{calls.push('durable');return {ok:true,records:payload.handoffs};}};
  const save=runInNewContext(`(${fn('updateExcerptHandoffStatuses')})`,context);
  const r=await save([{recordId:'a',payload:{sourceRow:2}}],{results:[{recordId:'a',ok:true,poetryPleaseItemId:'EXC'}]});
  assert.deepEqual(calls,['durable','sheet']);assert.equal(r.records[0].handoffStatus,'sent');
  assert.match(r.compatibilityWarning,/sheet compatibility update failed/);
  calls.length=0;context.syncWeaverRuntimeDb=async()=>{calls.push('durable');throw new Error('storage unavailable');};
  const failed=runInNewContext(`(${fn('updateExcerptHandoffStatuses')})`,context);
  await assert.rejects(failed([{recordId:'a',payload:{sourceRow:2}}],{}),/storage unavailable/);
  assert.deepEqual(calls,['durable']);
});

test('an HTTP success without receiver identity is an uncertain excerpt delivery', async () => {
  const send = runInNewContext(`(${fn('handoffApprovedExcerptsToPoetryPlease')})`, {
    buildPoetryPleaseExcerptRecord: r => r, poetryPleaseApiKey: 'test-key',
    poetryPleaseApiUrl: 'https://receiver.example/api', cleanSheetWhitespace: clean,
    fetch: async () => ({ ok: true, status: 200, json: async () => ({ ok: true, createdCount: 1 }) })
  });
  const result = await send([{ recordId: 'a', contentType: 'EXC' }]);
  assert.equal(result.ok, false); assert.equal(result.errorCount, 1);
  assert.match(result.results[0].error, /reconcile before retry/);
});

test('excerpt collection does not change ranking or candidate eligibility', async () => {
  const files = [{ id: 'with-excerpts', name: 'One' }, { id: 'without-excerpts', name: 'Two' }];
  const reviews = files.flatMap(file => ['a', 'b'].map(email => ({
    sourceFileId: file.id, reviewerEmail: email, rating: 'like', ratingSource: 'legacy_import',
    legacyScore: 6.95, excerptRecordIds: file.id === 'with-excerpts' ? ['excerpt-'+email] : []
  })));
  const build = runInNewContext(`(${fn('getVideoCurationCandidates')})`, {
    syncWeaverRuntimeDb: async action => ({ records: action === 'get_curation_reviews' ? reviews : [] }),
    videoFileMediaType, videoReviewProgress, PRIORITY_VIDEO_SETS: new Map([['event', { id: 'event', label: 'Event' }]]),
    cleanSheetWhitespace: clean, getVideoSetMediaType: async () => ({ sourceMediaType: 'FV' }),
    listVideoSetFiles: async () => files, isPriorityVideoFileIncluded: () => true,
    reconcileVideoReviews: (_files, rows) => rows, buildVideoCurationCandidateId: (set,id) => set+id,
    VIDEO_CURATION_SCORE_BY_RATING: { like: 7.5 }, VIDEO_CURATION_CANDIDATE_THRESHOLD: 7,
    getPriorityVideoReleaseStatus: () => '', requiresManualVideoIdentityReview: () => false, resolveVideoReviewSourceFileId: id => id,
    getFootageAssetConnections: async () => [], footageProvenanceForFile: () => [],
    poetryPleaseVideoIdForSourceFile: () => '', listKnownVideoAssets: () => [],
    resolvePublishingBookMeta: () => null, buildVideoReviewEvidence: () => []
  });
  const result = await build();
  assert.equal(result.candidates.length, 2);
  for (const c of result.candidates) {
    assert.equal(c.candidateScore, 6.95); assert.equal(c.excerptBonus, 0); assert.equal(c.isEligible, false);
  }
  assert.equal(result.candidates.find(c => c.sourceFileId === 'with-excerpts').excerptRecordIds.length, 2);
  const app = readFileSync(new URL('../public/app.js', import.meta.url), 'utf8');
  assert.doesNotMatch(app, /\+ excerpt bonus/);
  assert.match(app, /candidate\.excerptRecordIds\.length} excerpts/);
});
