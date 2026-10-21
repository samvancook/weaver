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

test("priority video progress and scoreboard remain visible from the gear", () => {
  assert.match(html, /<details[^>]*id="video-progress-management"[^>]*\bopen\b/);
  assert.match(html, /id="video-progress-set-rows"/);
  assert.match(html, /id="video-progress-reviewer-rows"/);
  assert.match(html, /id="video-curation-candidates"/);

  const gearHandler = app.match(/elements\.showGraphicsOps\?\.addEventListener\("click", \(\) => \{([\s\S]*?)\n\}\);/);
  assert.ok(gearHandler, "gear click handler must exist");
  assert.match(gearHandler[1], /currentGraphicsView = "ops"/);
  assert.match(gearHandler[1], /videoProgressManagement\.open = true/);
  assert.match(gearHandler[1], /loadPriorityVideoProgress\(\)/);
  assert.match(gearHandler[1], /loadVideoCurationCandidates\(\)/);
});

test("video scoreboard offers an event filter and ranks within event groups", () => {
  assert.match(html, /<select id="video-curation-event">/);
  assert.match(html, /<option value="">All events<\/option>/);
  assert.match(app, /videoCurationEvent\?\.addEventListener\("change", renderVideoCurationCandidates\)/);
  assert.match(app, /const groups = new Map\(\)/);
  assert.match(app, /candidates\.sort\(\(a, b\) => Number\(b\.candidateScore/);
});

test("scoreboard keeps reviewed videos below threshold and links every excerpt", () => {
  assert.match(server, /if \(!reviews\.length && !existingGate\) return null/);
  assert.match(server, /return candidate;\s*\}\)\.filter\(Boolean\)/);
  assert.match(server, /selectedExcerptRecordIds: candidate\.excerptRecordIds/);
  assert.match(server, /const gate = \{ \.\.\.candidate\.gate, selectedExcerptRecordIds: candidate\.excerptRecordIds, publishableAssetMediaType \}/);
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
  assert.match(server, /rawLegacyScore === "" \? NaN : Number\(rawLegacyScore\)/);
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
