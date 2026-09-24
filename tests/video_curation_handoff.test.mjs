import test from "node:test";
import assert from "node:assert/strict";
import { buildWeaverVideoImport, parsePoetryPleaseVideoImport, poetryPleaseVideoIdForSourceFile, reconcileVideoReviews, resolveCurrentVideoFileId } from "../video_curation_handoff.mjs";

const candidate = {
  candidateId: "weaver:video:lane-a:source-123",
  prioritySetId: "lane-a",
  sourceFileId: "source-123",
  sourceVideoUrl: "https://drive.google.com/file/d/source-123/view",
  sourceEvent: "BPL Charm City 2026",
  sourceEventLabel: "BPL Charm City 2026",
  author: "Author",
  poemTitle: "Poem",
  bookTitle: "Book",
  reviewCount: 2,
  excerptRecordIds: ["exc-1"],
  releaseStatus: "",
  publicationRestricted: false
};
const gate = {
  gateId: "gate-123",
  decision: "ready_for_poetry_please",
  publishableAssetUrl: "https://drive.google.com/file/d/final-456/view",
  selectedExcerptRecordIds: ["exc-1"]
};

test("video import is a single root record with stable source identity", () => {
  const record = buildWeaverVideoImport(candidate, gate);
  assert.equal(record.contentType, "VV");
  assert.equal(record.sourceRecordId, "weaver:video:source-123");
  assert.equal(record.sourceEventLabel, "BPL Charm City 2026");
  assert.equal(record.finalAssetUrl, gate.publishableAssetUrl);
  assert.deepEqual(record.selectedExcerptRecordIds, ["exc-1"]);
  assert.equal(record.records, undefined);
  assert.equal(buildWeaverVideoImport({ ...candidate, prioritySetId: "lane-b" }, gate).sourceRecordId, record.sourceRecordId);
});

test("video import carries individual reviews and excerpt IDs, not unapproved quote text", () => {
  const ratings = [{
    reviewId: "review-1", reviewerEmail: "reviewer@example.com", rating: "moved_me",
    notes: "Strong ending", excerptRecordIds: ["exc-1"], updatedAt: "2026-08-12T12:00:00Z"
  }];
  const record = buildWeaverVideoImport({ ...candidate, ratings }, gate);
  assert.deepEqual(record.reviews, ratings);
  assert.equal(record.excerpts, undefined);
  assert.deepEqual(record.selectedExcerptRecordIds, ["exc-1"]);
});

test("video import rejects raw footage and restricted or unready gates", () => {
  assert.throws(() => buildWeaverVideoImport(candidate, {
    ...gate, publishableAssetUrl: "https://drive.google.com/uc?id=source-123"
  }), /raw Weaver source/);
  assert.throws(() => buildWeaverVideoImport({ ...candidate, publicationRestricted: true }, gate), /publication-restricted/);
  assert.throws(() => buildWeaverVideoImport(candidate, { ...gate, decision: "send_to_editing" }), /ready-for-Poetry Please/);
});

test("replaced Drive files recover reviews only through a unique exact filename", () => {
  const files = [
    { id: "new-1080", name: "Rachel Mckibbens - Weather’s Here.mov" },
    { id: "another", name: "Other poem.mov" }
  ];
  assert.equal(resolveCurrentVideoFileId({
    sourceFileId: "old-720", sourceFileName: "Rachel Mckibbens - Weather’s Here.mov"
  }, files), "new-1080");
  assert.equal(resolveCurrentVideoFileId({
    sourceFileId: "new-1080", sourceFileName: "Stale name.mov"
  }, files), "new-1080");
  assert.equal(resolveCurrentVideoFileId({
    sourceFileId: "old-720", sourceFileName: "Rachel Mckibbens - Other poem.mov"
  }, files), "");
  assert.equal(resolveCurrentVideoFileId({
    sourceFileId: "old-720", sourceFileName: "Rachel Mckibbens - Weather’s Here.mov"
  }, [...files, { id: "duplicate", name: files[0].name }]), "");
});

test("Charm City Cam Audio label relocation preserves the reviewed video's identity", () => {
  const old = "[Vertical Version] George Abraham (Cam Audio) - Ode to Mennel Ibtissam Singing Hallelujah on The Voice (France) Translated in Arabic.mov";
  const current = "[Vertical Version] George Abraham - Ode to Mennel Ibtissam Singing Hallelujah on The Voice (France) Translated in Arabic (Cam Audio).mov";
  const files = [{ id: "1mF3LqQ8UNJOxmpse757wxvdBfuH7MF7Q", name: current }];
  assert.equal(resolveCurrentVideoFileId({ sourceFileId: "1YO78dpn_UobqRF9I9GCnVBAoCm5VtSZB", sourceFileName: old }, files), files[0].id);
  assert.equal(resolveCurrentVideoFileId({ sourceFileId: "old", sourceFileName: old }, [...files, { id: "duplicate", name: current }]), "");
});

test("a replaced review source stays the stable handoff identity", () => {
  const record = buildWeaverVideoImport({
    ...candidate,
    sourceFileId: "new-1080",
    sourceReviewFileId: "old-720"
  }, {
    ...gate,
    publishableAssetUrl: "https://drive.google.com/file/d/new-1080/view"
  });
  assert.equal(record.sourceRecordId, "weaver:video:old-720");
  assert.equal(record.sourceFileId, "old-720");
  assert.equal(record.sourceDriveFileId, "old-720");
  assert.equal(record.sourceVideoUrl, "https://drive.google.com/file/d/old-720/view");
  assert.equal(record.finalAssetUrl, "https://drive.google.com/file/d/new-1080/view");
  assert.equal(buildWeaverVideoImport({ ...candidate, sourceFileId: "new-1080" }, {
    ...gate, sourceReviewFileId: "old-720", publishableAssetUrl: record.finalAssetUrl
  }).sourceRecordId, record.sourceRecordId);
});

test("an updated reviewer counts once and retains excerpts from the archived review", () => {
  const files = [{ id: "new-1080", name: "Rachel Mckibbens - Weather’s Here.mov" }];
  const reviews = reconcileVideoReviews(files, [
    { sourceFileId: "old-720", sourceFileName: files[0].name, reviewerEmail: "sam@example.com", rating: "moved_me", excerptRecordIds: ["exc-old"] },
    { sourceFileId: "new-1080", sourceFileName: files[0].name, reviewerEmail: "sam@example.com", rating: "like", excerptRecordIds: ["exc-new"] }
  ]);
  assert.equal(reviews.length, 1);
  assert.equal(reviews[0].sourceFileId, "new-1080");
  assert.equal(reviews[0].rating, "like");
  assert.deepEqual(reviews[0].excerptRecordIds, ["exc-old", "exc-new"]);
  assert.deepEqual(reviews[0].originalSourceFileIds, ["old-720", "new-1080"]);
});

test("approved excerpts link only to a uniquely matched canonical video", () => {
  const gates = [{
    sourceFileId: "new-1080", sourceReviewFileId: "old-720",
    poetryPleaseHandoff: { canonicalVideoId: "WEAVER-VV-ABC" }
  }];
  assert.equal(poetryPleaseVideoIdForSourceFile("old-720", gates), "WEAVER-VV-ABC");
  assert.equal(poetryPleaseVideoIdForSourceFile("new-1080", gates), "WEAVER-VV-ABC");
  assert.equal(poetryPleaseVideoIdForSourceFile("other", gates), "");
  assert.equal(poetryPleaseVideoIdForSourceFile("old-720", [
    ...gates, { sourceFileId: "old-720", poetryPleaseHandoff: { canonicalVideoId: "OTHER" } }
  ]), "");
});

test("Poetry Please canonical response is stored from its documented fields", () => {
  const result = parsePoetryPleaseVideoImport(
    { ok: true, status: 200 },
    { ok: true, results: [{
      sourceRecordId: "weaver:video:source-123",
      status: "created",
      canonicalVideoId: "WEAVER-VV-ABC",
      canonicalVideoUrl: "/app?item=WEAVER-VV-ABC&type=VV",
      receivedReviewCount: 1,
      receivedSelectedExcerptIdCount: 1,
      finalAssetUrl: "https://poetryplease.org/video.mp4"
    }] },
    "weaver:video:source-123",
    "https://poetryplease.org/api",
    { reviewCount: 1, selectedExcerptIdCount: 1 }
  );
  assert.equal(result.ok, true);
  assert.equal(result.canonicalVideoId, "WEAVER-VV-ABC");
  assert.equal(result.canonicalVideoUrl, "https://poetryplease.org/app?item=WEAVER-VV-ABC&type=VV");
});

test("video import is not marked sent when Poetry Please drops reviews or excerpt links", () => {
  const result = parsePoetryPleaseVideoImport(
    { ok: true, status: 200 },
    { ok: true, results: [{
      sourceRecordId: "weaver:video:source-123", status: "updated",
      canonicalVideoId: "WEAVER-VV-ABC", canonicalVideoUrl: "/app?item=WEAVER-VV-ABC&type=VV"
    }] },
    "weaver:video:source-123", "https://poetryplease.org/api",
    { reviewCount: 1, selectedExcerptIdCount: 1 }
  );
  assert.equal(result.ok, false);
  assert.equal(result.status, "failed");
  assert.match(result.error, /did not confirm receipt/);
});
