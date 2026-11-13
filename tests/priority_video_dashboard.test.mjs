import {videoQualitySummary} from '../video_pipeline.mjs';
import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { runInNewContext } from "node:vm";

const html = readFileSync(new URL("../public/index.html", import.meta.url), "utf8");
const app = readFileSync(new URL("../public/app.js", import.meta.url), "utf8");
const server = readFileSync(new URL("../server.mjs", import.meta.url), "utf8");

test("graphics queue rows build without video-intake scope errors", () => {
  const source = server.match(/function buildGraphicsRequestRecordFromQueueRow\([\s\S]*?\n}\n\nfunction buildGraphicsHandoffLedgerRequest/);
  assert.ok(source);
  const build = runInNewContext(`(${source[0].replace(/\n\nfunction buildGraphicsHandoffLedgerRequest$/, "")})`, {
    cleanSheetWhitespace: value => String(value ?? "").trim(),
    resolveGraphicsAuthor: value => value,
    parseIntakeMetadataFromNotes: () => ({ curationRating: "like", curationLegacyScore: 8, curationNotes: "Strong" }),
    buildWeaverGraphicsRequestId: () => "weaver:qi:existing",
  });
  const row = Array(18).fill("");
  row[0] = "Leslie Kaup";
  row[1] = "Selected poem";
  row[2] = "Book";
  row[3] = "Verified quote";
  row[8] = "submission-aNT7jMmPfZ7eoyPxOsYr";
  const record = build(row, 0);
  assert.equal(record.graphicsRequestId, "weaver:qi:existing");
  assert.equal(record.quoteText, "Verified quote");
  assert.equal(record.curationRating, "like");
  assert.equal(record.curationLegacyScore, 8);
  assert.equal(record.curationNotes, "Strong");
});

test("Videos opens a dedicated workspace without loading graphics queues", () => {
  assert.match(html, /id="show-video-module"/);
  const start=html.indexOf('id="video-module"'), end=html.indexOf('id="graphics-module"');
  assert.ok(start>0 && end>start);
  for(const id of ['video-progress-management','video-curation-candidates','video-progress-set-rows','video-excerpt-delivery']) {
    assert.equal(html.split(`id="${id}"`).length-1,1);
    assert.ok(html.indexOf(`id="${id}"`)>start && html.indexOf(`id="${id}"`)<end);
  }
  const calls=[];
  const open=runInNewContext(`(${app.match(/function openVideoWorkspace\([^]*?\n}\n/)[0]})`,{
    setActiveModule:name=>calls.push(name),elements:{videoProgressManagement:{}},
    loadVideoCurationCandidates:()=>calls.push('candidates'),loadVideoExcerptDelivery:()=>calls.push('excerpts')
  });open();assert.deepEqual(calls,['videos','candidates','excerpts']);
});

test("video scoreboard offers an event filter and ranks within event groups", () => {
  assert.match(html, /<select id="video-curation-event">/);
  assert.match(html, /<option value="">All events<\/option>/);
  assert.match(app, /videoCurationEvent\?\.addEventListener\("change", \(\) => \{ renderVideoCurationCandidates\(\); renderVideoExcerptDelivery\(\); \}\)/);
  assert.match(app, /const groups = new Map\(\)/);
  assert.match(app, /candidates\.sort\(\(a, b\) => Number\(b\.candidateScore/);
});

test("scoreboard keeps below-threshold reviews and preserves the saved excerpt selection", () => {
  assert.match(server, /if \(!reviews\.length && !existingGate\) return null/);
  assert.match(server, /return candidate;\s*\}\)\.filter\(Boolean\)/);
  assert.match(server, /selectedExcerptRecordIds: candidate\.excerptRecordIds/);
  assert.match(server, /const gate = \{ \.\.\.candidate\.gate, publishableAssetMediaType \}/);
  assert.doesNotMatch(app, /data-excerpt-id/);
  assert.match(app, /candidate\.ratings/);
  assert.match(app, /candidate\.baseScore/);
});

test("LABC intake keeps the four supplied files review-only until approval evidence is linked", () => {
  const set = server.match(/\["pages-matam-san-diego-2026-labc", \{([\s\S]*?)\n  \}\]/);
  assert.ok(set);
  assert.match(set[1], /eventName: "PAGES Matam in San Diego - 2026"/);
  assert.match(set[1], /bookTitle: "Living at Baggage Claim"/);
  assert.match(set[1], /requireUsageApprovalEvidence: true/);
  for (const fileId of [
    "14NdOdzkV1F4dScR2QYk8k20WTAlyJF2g",
    "1eoqefV0zfkzR9fEZ8Zr4HFY5rsGVw-FA",
    "1Te0zeh7tmd820BBoUw2HOcDCbkBMjGwH",
    "1MwOqcF4HWhB2ukcs58mzbbRU2xunipTO"
  ]) assert.ok(set[1].includes(fileId));
  assert.match(server, /if \(candidate\.publicationRestricted\) throw new Error\("A publication-restricted source cannot be sent to Poetry Please\."\)/);
});

test("empty imported numeric scores are not displayed or averaged as zero", () => {
  assert.equal(videoQualitySummary([{reviewerEmail:'a',ratingSource:'legacy_import',legacyScore:'',rating:'like'}]).candidateScore,7.5);
  assert.equal(videoQualitySummary([{reviewerEmail:'a',ratingSource:'legacy_import',legacyScore:0,rating:'like'}]).candidateScore,0);
  assert.match(app, /rating\.legacyScore !== "" && Number\.isFinite\(Number\(rating\.legacyScore\)\)/);
});

test("parallel admin loads share token acquisition and progress uses linked excerpt IDs", () => {
  assert.match(app, /if \(googleAdminTokenRequest\) return googleAdminTokenRequest/);
  assert.match(app, /googleAdminTokenRequest = new Promise/);
  assert.match(app, /googleAdminTokenRequest = null/);
  assert.match(app, /Progress request timed out/);
  const progress = server.match(/async function loadPriorityVideoProgress\(\) \{([\s\S]*?)\n\}\n\nfunction buildVideoCurationCandidateId/);
  assert.ok(progress);
  assert.doesNotMatch(progress[1], /get_excerpt_records/);
  assert.match(progress[1], /const setExcerptIds = new Set/);
});
