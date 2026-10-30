import test from "node:test";
import assert from "node:assert/strict";
import { assertExistingCanonicalReceipt, buildReviewerProgressExport, buildWeaverVideoImport, classifyDriveFolderPath, listKnownVideoAssets, parsePoetryPleaseVideoImport, parsePoetryPleaseVideoReceipt, poetryPleaseVideoIdForSourceFile, reconcileVideoReviews, requiresManualVideoIdentityReview, resolveCurrentVideoFileId, resolveVideoReviewSourceFileId } from "../video_curation_handoff.mjs";

test("receipt refresh accepts only the same event, identity, retained context and verified hosted copy", () => {
  const gate = { sourceEvent: "event", sourceEventLabel: "Event", selectedExcerptRecordIds: ["excerpt"],
    poetryPleaseHandoff: { sourceRecordId: "source", canonicalVideoId: "canonical", receivedReviewCount: 11, receivedSelectedExcerptIdCount: 8 } };
  const receipt = { ok: true, received: true, sourceRecordId: "source", canonicalVideoId: "canonical",
    event: { sourceEvent: "event", sourceEventLabel: "Event" }, review: { reviewCount: 11, selectedExcerptIdCount: 8 },
    media: { durableCopyStatus: "verified", storageMode: "durable", mediaStatus: "verified", reviewOnly: true,
      durableByteCount: 1447300329, durableChecksum: "checksum", videoUrl: "https://firebasestorage.googleapis.com/v0/b/pp/o/final" } };
  const parsed = parsePoetryPleaseVideoReceipt(receipt, gate);
  assert.equal(parsed.finalAssetUrl, receipt.media.videoUrl);
  assert.equal(parsed.durableByteCount, 1447300329);
  assert.equal(parsed.reviewOnly, true);
  assert.equal(parsed.storedReviewCount, 11);
  for (const changed of [
    { ...receipt, canonicalVideoId: "other" },
    { ...receipt, event: { ...receipt.event, sourceEvent: "other" } },
    { ...receipt, review: { ...receipt.review, reviewCount: 9 } },
    { ...receipt, media: { ...receipt.media, durableByteCount: 0 } },
    { ...receipt, media: { ...receipt.media, videoUrl: "https://example.com/unverified" } },
  ]) assert.throws(() => parsePoetryPleaseVideoReceipt(changed, gate));
  const pending = parsePoetryPleaseVideoReceipt({ ...receipt, media: { ...receipt.media, durableCopyStatus: "copying" } }, gate);
  assert.equal(Object.hasOwn(pending, "finalAssetUrl"), false);
});

test("multiple historical files resolve through a unique earlier canonical receipt", () => {
  const earlier = { prioritySetId: "event", sourceFileId: "intermediate", sourceReviewFileId: "reviewed-720",
    poetryPleaseHandoff: { status: "sent_to_poetry_please", sourceRecordId: "weaver:video:reviewed-720", canonicalVideoId: "canonical" } };
  const history = ["reviewed-720", "intermediate"];
  assert.equal(resolveVideoReviewSourceFileId("final-2160", history, [earlier], { prioritySetId: "event" }), "reviewed-720");
  assert.equal(resolveVideoReviewSourceFileId("final-2160", history, [], {}), "");
  assert.equal(resolveVideoReviewSourceFileId("final-2160", history, [earlier], { prioritySetId: "other-event" }), "");
  assert.equal(resolveVideoReviewSourceFileId("final-2160", history, [earlier, {
    ...earlier, poetryPleaseHandoff: { ...earlier.poetryPleaseHandoff, canonicalVideoId: "conflicting-canonical" },
  }], {}), "");
  assert.equal(resolveVideoReviewSourceFileId("final-2160", ["reviewed-720"], [], {}), "reviewed-720");
});

test("multiple resolutions and formats retain one canonical source identity", () => {
  const gates = [
    { prioritySetId: "event", sourceFileId: "vertical-1080", sourceReviewFileId: "review-720",
      poetryPleaseHandoff: { status: "sent_to_poetry_please", sourceRecordId: "weaver:video:review-720", canonicalVideoId: "canonical" } },
    { prioritySetId: "event", sourceFileId: "horizontal-mp4", sourceReviewFileId: "review-720",
      poetryPleaseHandoff: { status: "sent_to_poetry_please", sourceRecordId: "weaver:video:review-720", canonicalVideoId: "canonical" } },
  ];
  for (const fileId of ["vertical-2160", "square-webm"]) {
    assert.equal(resolveVideoReviewSourceFileId(fileId,
      ["review-720", "vertical-1080", "horizontal-mp4"], gates, { prioritySetId: "event" }), "review-720");
  }
  assert.equal(poetryPleaseVideoIdForSourceFile("review-720", gates), "canonical");
  const assets = listKnownVideoAssets({ id: "square-webm" },
    ["review-720", "vertical-1080", "horizontal-mp4"],
    [...gates, { sourceFileId: "unrelated", sourceReviewFileId: "someone-else",
      publishableAssetUrl: "https://drive.google.com/file/d/unrelated-final/view" }], "review-720");
  assert.deepEqual(assets.map(asset => asset.fileId),
    ["square-webm", "review-720", "vertical-1080", "horizontal-mp4"]);
  assert.equal(assets.some(asset => asset.fileId === "unrelated-final"), false);
  const receipt = { ok: true, received: true, sourceRecordId: "weaver:video:review-720",
    canonicalVideoId: "canonical", event: { sourceEvent: "BPL Charm City 2026" } };
  assert.doesNotThrow(() => assertExistingCanonicalReceipt(receipt, receipt.sourceRecordId, "canonical", "BPL Charm City 2026"));
  for (const changed of [
    { ...receipt, canonicalVideoId: "another" },
    { ...receipt, ambiguousVideoMatch: true },
    { ...receipt, event: { sourceEvent: "another event" } },
    { ...receipt, received: false },
  ]) assert.throws(() => assertExistingCanonicalReceipt(changed, receipt.sourceRecordId, "canonical", "BPL Charm City 2026"));
});

test("a failed legacy attempt preserves its explicit original source identity", () => {
  const old = { prioritySetId: "event", sourceFileId: "intermediate", sourceReviewFileId: "original",
    videoRecordId: "weaver:video:original", poetryPleaseHandoff: { status: "failed", ok: false,
      sourceRecordId: "weaver:video:original", error: "The operation was aborted due to timeout" } };
  const history = ["original", "intermediate"];
  const current = { prioritySetId: "event" };
  assert.equal(resolveVideoReviewSourceFileId("new-final", history, [old], current), "original");
  for (const bad of [
    { ...old, prioritySetId: "another-event" },
    { ...old, sourceReviewFileId: "" },
    { ...old, videoRecordId: "weaver:video:other" },
    { ...old, poetryPleaseHandoff: { ...old.poetryPleaseHandoff, sourceRecordId: "weaver:video:other" } },
    { ...old, poetryPleaseHandoff: { ...old.poetryPleaseHandoff, ok: true } },
    { ...old, poetryPleaseHandoff: { ...old.poetryPleaseHandoff, canonicalVideoId: "unexpected" } },
    { ...old, poetryPleaseHandoff: {} },
  ]) assert.equal(resolveVideoReviewSourceFileId("new-final", history, [bad], current), "");
  assert.equal(resolveVideoReviewSourceFileId("new-final", history, [old], {}), "");
  const conflict = { ...old, sourceFileId: "original", sourceReviewFileId: "intermediate",
    videoRecordId: "weaver:video:intermediate", poetryPleaseHandoff: { ...old.poetryPleaseHandoff,
      sourceRecordId: "weaver:video:intermediate" } };
  assert.equal(resolveVideoReviewSourceFileId("new-final", history, [old, conflict], current), "");
  const sentConflict = { ...conflict, poetryPleaseHandoff: { status: "sent_to_poetry_please",
    sourceRecordId: "weaver:video:intermediate", canonicalVideoId: "other-canonical" } };
  assert.equal(resolveVideoReviewSourceFileId("new-final", history, [old, sentConflict], current), "");
});

test("receiver validation errors stay in finalizable held states", () => {
  const sourceRecordId = "weaver:video:original";
  for (const [status, expected] of [["blocked", "blocked"], ["error", "failed"], ["unexpected", "failed"]]) {
    const result = parsePoetryPleaseVideoImport({ ok: false, status: 409 }, {
      ok: false, results: [{ sourceRecordId, status, error: "receiver_validation" }],
    }, sourceRecordId, "https://poetryplease.org/api");
    assert.equal(result.ok, false);
    assert.equal(result.status, expected);
    assert.equal(result.error, "receiver_validation");
  }
});

test("accepted review-only receipts retain copy state without claiming publication", () => {
  const result = parsePoetryPleaseVideoImport({ ok: true, status: 200 }, { ok: true, results: [{
    sourceRecordId: "source", status: "updated", canonicalVideoId: "canonical",
    canonicalVideoUrl: "/app?item=canonical", receivedReviewCount: 9,
    receivedSelectedExcerptIdCount: 8, storedReviewCount: 11, storedSelectedExcerptIdCount: 8,
    durableCopyStatus: "pending", durableCopyQueued: true, reviewOnly: true,
    storageMode: "external", mediaStatus: "review_ready",
  }] }, "source", "https://poetryplease.org/api", { reviewCount: 9, selectedExcerptIdCount: 8 });
  assert.equal(result.ok, true);
  assert.equal(result.storedReviewCount, 11);
  assert.equal(result.receivedReviewCount, 9);
  assert.equal(result.durableCopyStatus, "pending");
  assert.equal(result.reviewOnly, true);
});

test("reviewer progress export includes unstarted reviewers and reconciles replacement files", () => {
  const files = [
    { id: "new-file", name: "Poem A.mov", webViewLink: "https://drive.google.com/file/d/new-file/view" },
    { id: "other-file", name: "Poem B.mov" }
  ];
  const reviews = [{ prioritySetId: "lane-a", reviewerEmail: "Reviewer@Example.com", sourceFileId: "old-file", sourceFileName: "Poem A.mov", rating: "like", excerptRecordIds: ["exc-1", "exc-1"], updatedAt: "2026-08-12T12:00:00Z" }];
  const input = [{ set: { id: "lane-a", label: "Lane A" }, files, reviews }];
  const report = buildReviewerProgressExport("REVIEWER@example.com", input);
  assert.deepEqual(report.totals, { prioritySets: 1, videos: 2, reviewed: 1, remaining: 1 });
  assert.deepEqual(report.sets[0].videos.map(video => video.reviewStatus), ["reviewed", "not_reviewed"]);
  assert.equal(report.sets[0].videos[0].sourceFileId, "new-file");
  assert.equal(report.sets[0].videos[0].excerptCount, 1);
  assert.equal(report.sets[0].videos[0].rating, "like");
  assert.equal(report.sets[0].videos[1].videoUrl, "https://drive.google.com/file/d/other-file/view");
  assert.equal(report.sets[0].videos[1].rating, null);
  assert.equal(buildReviewerProgressExport("new@example.com", input).totals.remaining, 2);
});

const candidate = {
  candidateId: "weaver:video:lane-a:source-123",
  prioritySetId: "lane-a",
  sourceFileId: "source-123",
  sourceMediaType: "FV",
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
  publishableAssetMediaType: "FV",
  selectedExcerptRecordIds: ["exc-1"]
};

const footageRef={inventoryEventId:'inventory-39a607f0dd29e65a111e13892bf0b17b',inventoryEventName:'PAGES footage from Rudy Francisco 2026',
 sourceSpreadsheetId:'1npv-dpP22MV-cRs-DRjOimjbWIxf-xvbun5SdP6dl7U',sourceRow:6871,
 familyId:'footage-family-3e64fa77f2dd9229f0b1d6c41243324b',originalMediaFileId:'1pHdxNXRNgbdzRecfUYYsvNt1VyZS7Wws',
 finishedVideoFileId:'1Te0zeh7tmd820BBoUw2HOcDCbkBMjGwH',curationFileId:'1Te0zeh7tmd820BBoUw2HOcDCbkBMjGwH'};
test('the PP handoff requires acknowledgement of exact inventory and OM/FV references',()=>{
 const payload=buildWeaverVideoImport({...candidate,footageConnections:[footageRef]},gate);
 assert.equal(payload.footageContextVersion,1);assert.deepEqual(payload.footageConnections,[footageRef]);
 assert.equal(payload.sourceRecordId,buildWeaverVideoImport(candidate,gate).sourceRecordId);
 const body={ok:true,results:[{sourceRecordId:payload.sourceRecordId,status:'created',canonicalVideoId:'canonical',canonicalVideoUrl:'/app?item=canonical',receivedReviewCount:0,receivedSelectedExcerptIdCount:0,
  receivedFootageConnections:[footageRef],receivedFootageConnectionCount:1,storedFootageConnectionCount:1}]};
 const expected={footageConnections:[footageRef]};
 assert.equal(parsePoetryPleaseVideoImport({ok:true},body,payload.sourceRecordId,'https://poetryplease.org',expected).ok,true);
 const wrong=structuredClone(body);wrong.results[0].receivedFootageConnections[0].originalMediaFileId='wrong_original_file_123456';
 assert.equal(parsePoetryPleaseVideoImport({ok:true},wrong,payload.sourceRecordId,'https://poetryplease.org',expected).ok,false);
 const dropped=structuredClone(body);delete dropped.results[0].receivedFootageConnections;delete dropped.results[0].receivedFootageConnectionCount;
 assert.match(parsePoetryPleaseVideoImport({ok:true},dropped,payload.sourceRecordId,'https://poetryplease.org',expected).error,/OM\/FV references/);
 const legacyGate={...gate,footageConnections:[footageRef]};
 assert.deepEqual(buildWeaverVideoImport(candidate,legacyGate).footageConnections,[footageRef]);
});
test('a later PP receipt cannot lose acknowledged footage references',()=>{
 const g={sourceEvent:'event',sourceEventLabel:'Event',selectedExcerptRecordIds:[],poetryPleaseHandoff:{sourceRecordId:'source',canonicalVideoId:'canonical',receivedFootageConnections:[footageRef]}};
 const r={ok:true,received:true,sourceRecordId:'source',canonicalVideoId:'canonical',event:{sourceEvent:'event',sourceEventLabel:'Event'},review:{reviewCount:0,selectedExcerptIdCount:0},media:{durableCopyStatus:'pending'},footage:{schemaVersion:1,connections:[footageRef]}};
 assert.equal(parsePoetryPleaseVideoReceipt(r,g).receivedFootageConnectionCount,1);
 assert.throws(()=>parsePoetryPleaseVideoReceipt({...r,footage:{connections:[]}},g),/missing previously received footage/);
});

test("a replacement file retains the saved gate's excerpt associations", () => {
  const record = buildWeaverVideoImport(candidate, {
    ...gate, selectedExcerptRecordIds: ["archived-excerpt", "exc-1"],
  });
  assert.deepEqual(record.selectedExcerptRecordIds, ["archived-excerpt", "exc-1"]);
});

test("video import is a single root record with stable source identity", () => {
  const record = buildWeaverVideoImport(candidate, gate);
  assert.equal(record.contentType, "VV");
  assert.equal(record.sourceRecordId, "weaver:video:source-123");
  assert.equal(record.sourceEventLabel, "BPL Charm City 2026");
  assert.equal(record.releaseCatalog, "BPL Charm City 2026");
  assert.equal(record.book, "Book");
  assert.equal(record.finalAssetUrl, gate.publishableAssetUrl);
  assert.deepEqual(record.selectedExcerptRecordIds, ["exc-1"]);
  assert.equal(record.records, undefined);
  assert.equal(buildWeaverVideoImport({ ...candidate, prioritySetId: "lane-b" }, gate).sourceRecordId, record.sourceRecordId);
});

test("Charm City video catalog overrides a stale book catalog without changing other events", () => {
  const charmCity = buildWeaverVideoImport({ ...candidate, releaseCatalog: "Spring 2022" }, gate);
  assert.equal(charmCity.releaseCatalog, "BPL Charm City 2026");
  assert.equal(charmCity.sourceEvent, "BPL Charm City 2026");
  assert.equal(charmCity.book, "Book");
  const otherEvent = buildWeaverVideoImport({
    ...candidate, sourceEvent: "Other event", sourceEventLabel: "Other event", releaseCatalog: "Fall 2026"
  }, gate);
  assert.equal(otherEvent.releaseCatalog, "Fall 2026");
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
  assert.throws(() => buildWeaverVideoImport({ ...candidate, sourceMediaType: "OM" }, {
    ...gate, publishableAssetUrl: "https://drive.google.com/uc?id=source-123"
  }), /raw Weaver source/);
  assert.throws(() => buildWeaverVideoImport(candidate, { ...gate, publishableAssetMediaType: "unknown" }), /verified as Finished Video/);
  assert.throws(() => buildWeaverVideoImport({ ...candidate, publicationRestricted: true }, gate), /publication-restricted/);
  assert.throws(() => buildWeaverVideoImport(candidate, { ...gate, decision: "send_to_editing" }), /ready-for-Poetry Please/);
});

test("direct FV handoff needs identity review, while folder conflicts stay unknown", () => {
  assert.throws(() => buildWeaverVideoImport(candidate, {
    ...gate, publishableAssetUrl: candidate.sourceVideoUrl
  }), /canonical identity review/);
  assert.equal(requiresManualVideoIdentityReview(candidate, candidate.sourceVideoUrl), true);
  assert.equal(requiresManualVideoIdentityReview(candidate, gate.publishableAssetUrl), false);
  assert.equal(requiresManualVideoIdentityReview({ ...candidate, sourceMediaType: "OM" }, candidate.sourceVideoUrl), false);
  assert.equal(requiresManualVideoIdentityReview({ ...candidate, sourceReviewFileId: "review-720" }, candidate.sourceVideoUrl), false);
  assert.equal(classifyDriveFolderPath([
    { id: "child", name: "FV - Vertical Versions - BPL Charm City 2026" },
    { id: "parent", name: "Vertical Video | 9:16" }
  ]).sourceMediaType, "FV");
  assert.equal(classifyDriveFolderPath([
    { id: "child", name: "Camera Y" },
    { id: "parent", name: "Button Footage - Original Media (OM)" }
  ]).sourceMediaType, "OM");
  assert.equal(classifyDriveFolderPath([
    { id: "child", name: "OM for completed longform projects" },
    { id: "parent", name: "FV BTV VJ" }
  ]).sourceMediaType, "unknown");
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

test("Poetry Please catalog warnings fail an otherwise successful video import", () => {
  const result = parsePoetryPleaseVideoImport(
    { ok: true, status: 200 },
    { ok: true, results: [{
      sourceRecordId: "weaver:video:source-123", status: "updated",
      canonicalVideoId: "WEAVER-VV-ABC", canonicalVideoUrl: "/app?item=WEAVER-VV-ABC&type=VV",
      receivedReviewCount: 1, receivedSelectedExcerptIdCount: 1,
      warnings: ["missing_release_catalog"]
    }] },
    "weaver:video:source-123", "https://poetryplease.org/api",
    { reviewCount: 1, selectedExcerptIdCount: 1 }
  );
  assert.equal(result.ok, false);
  assert.equal(result.status, "failed");
  assert.deepEqual(result.warnings, ["missing_release_catalog"]);
  assert.match(result.error, /missing_release_catalog/);
});

test("top-level Poetry Please warnings also fail a video import", () => {
  const result = parsePoetryPleaseVideoImport(
    { ok: true, status: 200 },
    { ok: true, warnings: ["missing_release_catalog"], results: [{
      sourceRecordId: "weaver:video:source-123", status: "updated",
      canonicalVideoId: "WEAVER-VV-ABC", canonicalVideoUrl: "/app?item=WEAVER-VV-ABC&type=VV",
      receivedReviewCount: 1, receivedSelectedExcerptIdCount: 1
    }] },
    "weaver:video:source-123", "https://poetryplease.org/api",
    { reviewCount: 1, selectedExcerptIdCount: 1 }
  );
  assert.equal(result.ok, false);
  assert.deepEqual(result.warnings, ["missing_release_catalog"]);
});
