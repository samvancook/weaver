import test from "node:test";
import assert from "node:assert/strict";
import { buildWeaverVideoImport, parsePoetryPleaseVideoImport } from "../video_curation_handoff.mjs";

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

test("video import rejects raw footage and restricted or unready gates", () => {
  assert.throws(() => buildWeaverVideoImport(candidate, {
    ...gate, publishableAssetUrl: "https://drive.google.com/uc?id=source-123"
  }), /raw Weaver source/);
  assert.throws(() => buildWeaverVideoImport({ ...candidate, publicationRestricted: true }, gate), /publication-restricted/);
  assert.throws(() => buildWeaverVideoImport(candidate, { ...gate, decision: "send_to_editing" }), /ready-for-Poetry Please/);
});

test("Poetry Please canonical response is stored from its documented fields", () => {
  const result = parsePoetryPleaseVideoImport(
    { ok: true, status: 200 },
    { ok: true, results: [{
      sourceRecordId: "weaver:video:source-123",
      status: "created",
      canonicalVideoId: "WEAVER-VV-ABC",
      canonicalVideoUrl: "/app?item=WEAVER-VV-ABC&type=VV",
      finalAssetUrl: "https://poetryplease.org/video.mp4"
    }] },
    "weaver:video:source-123",
    "https://poetryplease.org/api"
  );
  assert.equal(result.ok, true);
  assert.equal(result.canonicalVideoId, "WEAVER-VV-ABC");
  assert.equal(result.canonicalVideoUrl, "https://poetryplease.org/app?item=WEAVER-VV-ABC&type=VV");
});
