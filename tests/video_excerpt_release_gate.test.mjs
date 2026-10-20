import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { runInNewContext } from "node:vm";

const source = readFileSync(new URL("../server.mjs", import.meta.url), "utf8");
const gateStart = source.indexOf("function approvedVideoFileIdsFromGates(");
const gateEnd = source.indexOf("async function getApprovedVideoFileIds(", gateStart);
const releaseStart = source.indexOf("function isVideoExcerptReleased(");
const releaseEnd = source.indexOf("async function getPendingRecordsFromSheets(", releaseStart);
assert.ok(gateStart >= 0 && gateEnd > gateStart && releaseStart >= 0 && releaseEnd > releaseStart);

const { approvedVideoFileIdsFromGates, isVideoExcerptReleased } = runInNewContext(
  `${source.slice(gateStart, gateEnd)}\n${source.slice(releaseStart, releaseEnd)}\n({ approvedVideoFileIdsFromGates, isVideoExcerptReleased })`,
  { cleanSheetWhitespace: value => String(value || "").trim() }
);

test("only approved video gates release their exact source files", () => {
  const allowed = approvedVideoFileIdsFromGates([
    { decision: "ready_for_poetry_please", sourceFileId: "new-file", sourceReviewFileId: "old-file" },
    { decision: "send_to_editing", sourceFileId: "editing-file" },
    { decision: "hold", sourceFileId: "held-file" }
  ]);
  assert.equal(isVideoExcerptReleased({ intakeMode: "video", sourceVideoFileId: "new-file" }, allowed), true);
  assert.equal(isVideoExcerptReleased({ intakeMode: "video", sourceVideoFileId: "old-file" }, allowed), true);
  assert.equal(isVideoExcerptReleased({ intakeMode: "video", sourceVideoFileId: "editing-file" }, allowed), false);
  assert.equal(isVideoExcerptReleased({ intakeMode: "video", sourceVideoFileId: "" }, allowed), false);
  assert.equal(isVideoExcerptReleased({ intakeMode: "book", sourceVideoFileId: "" }, allowed), true);
});

const handoffStart = source.indexOf("function buildPoetryPleaseExcerptRecord(");
const handoffEnd = source.indexOf("async function handoffApprovedExcerptsToPoetryPlease(", handoffStart);
assert.ok(handoffStart >= 0 && handoffEnd > handoffStart);
const { buildPoetryPleaseExcerptRecord } = runInNewContext(
  `${source.slice(handoffStart, handoffEnd)}\n({ buildPoetryPleaseExcerptRecord })`,
  {
    cleanSheetWhitespace: value => String(value || "").trim(),
    normalizeExcerptTransferText: value => String(value || "").trim(),
    normalizeExcerptContentType: value => String(value || "").toUpperCase(),
    extractGoogleDriveFileId: value => String(value || "").match(/\/d\/([^/]+)/)?.[1] || "",
    normalizeCurationRating: () => "",
    normalizeOptionalCurationScore: () => null,
    canonicalFullPoemSourceRecordId: () => "",
    resolveExcerptPoemKey: () => ({ poemKey: null, poemKeySource: null })
  }
);

test("video-linked excerpts may omit book metadata without weakening book imports", () => {
  const base = {
    contentType: "EXC",
    recordId: "weaver:row-1",
    sourceRecordId: "weaver:source-1",
    author: "Poet",
    poemTitle: "Poem",
    excerpt: "Selected lines"
  };
  const video = buildPoetryPleaseExcerptRecord({
    ...base,
    sourceVideoUrl: "https://drive.google.com/file/d/video-file/view",
    sourceEvent: "BPL Charm City 2026"
  });
  assert.equal(video.invalid, undefined);
  assert.equal(video.sourceRecordId, base.sourceRecordId);
  assert.equal(video.sourceVideoFileId, "video-file");
  assert.equal(video.bookShortener, "");
  assert.equal(video.releaseCatalog, "");

  const book = buildPoetryPleaseExcerptRecord(base);
  assert.equal(book.invalid, true);
  assert.match(book.error, /bookShortener, releaseCatalog/);
  const missingSource = buildPoetryPleaseExcerptRecord({ ...base, sourceRecordId: "", sourceVideoFileId: "video-file" });
  assert.equal(missingSource.invalid, true);
  assert.match(missingSource.error, /sourceRecordId/);
});
