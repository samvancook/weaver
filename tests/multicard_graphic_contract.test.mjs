import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import { createHash } from "node:crypto";

const source = readFileSync(new URL("../server.mjs", import.meta.url), "utf8");
const assets = [1, 2, 3].map(index => ({
  index, role: "card", assetUrl: `https://example.com/card-${index}.png`,
  assetPreviewUrl: `https://example.com/preview-${index}.png`,
  assetFileId: `synthetic-${index}`, fileName: `card-${index}.png`
}));
function loadFunctions(names, overrides = {}) {
  const context = vm.createContext({
    cleanSheetWhitespace: value => String(value || "").trim(),
    resolvePublishingBookMeta: () => ({}),
    extractGoogleDriveFileId: () => "",
    inferPoetryPleaseContentType: record => record.contentType,
    slugToken: value => value,
    weaverPublicBaseUrl: "https://weaver.example.com",
    poetryPleaseApiKey: "synthetic-key", poetryPleaseApiUrl: "https://pp.example.com",
    AbortSignal, console: { log() {} }, ...overrides
  });
  for (const name of names) {
    const match = source.match(new RegExp(`(?:async )?function ${name}\\([^]*?^}`, "m"));
    assert.ok(match, name);
    vm.runInContext(match[0], context);
  }
  return context;
}

test("legacy request-wrapped content identity validates without accepting other content", async () => {
  const requestId = "pig:record-TGIT-FP-JOHN-MCPHEE";
  const project = {
    id: "saved-project", pigProjectId: "saved-project",
    sourceIdentity: { graphicsRequestId: requestId },
    selectedRecord: { contentId: "TGIT-FP-JOHN-MCPHEE", imageId: "TGIT-FP-JOHN-MCPHEE",
      graphicsRequestId: requestId, quoteText: "Synthetic source text" },
    multiCardDraft: { cards: [{ titleText: "Card 1" }, { titleText: "Card 2" }] }
  };
  const completion = { completionId: "synthetic-completion", graphicsRequestId: requestId,
    sourceRecordId: requestId, quoteText: "Synthetic source text", pigProjectId: project.id,
    editableProjectFileId: "saved-json", editableProjectUrl: "https://example.com/saved.json",
    editableProjectValidationError: { message: "old conflict" } };
  const context = loadFunctions([
    "normalizeEditableIdentityText", "cleanEditableIdentityValue", "editableIdentityTextHash",
    "getEditableIdentityValues", "validateCompletionEditableProjectIdentity", "validateEditableProjectForRework"
  ], { createHash, fetchDriveFileResponse: async () => ({ ok: true, json: async () => project }),
    syncWeaverRuntimeDb: async () => ({ record: {
      ...completion, id: completion.completionId, pigProjectId: "", editableProjectFileId: "",
      editableProjectUrl: "", candidatePigProjectId: project.id,
      candidateEditableProjectFileId: "saved-json", candidateEditableProjectUrl: completion.editableProjectUrl
    } }) });
  const verified = await context.validateCompletionEditableProjectIdentity(completion);
  assert.equal(verified.editableProjectValidationStatus, "verified");
  assert.equal(Object.keys(verified.editableProjectValidationError).length, 0);
  const rework = await context.validateEditableProjectForRework({
    pigCompletionId: completion.completionId, graphicsRequestId: requestId
  });
  assert.equal(rework.editableProjectAvailable, true);
  assert.equal(rework.editableProjectFileId, "saved-json");
  project.selectedRecord.contentId = "OTHER-POEM";
  project.selectedRecord.imageId = "OTHER-POEM";
  await assert.rejects(context.validateCompletionEditableProjectIdentity(completion), /content identity conflicts/);
  project.selectedRecord.contentId = "TGIT-FP-JOHN-MCPHEE";
  project.selectedRecord.imageId = "TGIT-FP-JOHN-MCPHEE";
  project.selectedRecord.quoteText = "Different source text";
  await assert.rejects(context.validateCompletionEditableProjectIdentity(completion), /text conflicts/);
  project.sourceIdentity.graphicsRequestId = "different-request";
  await assert.rejects(context.validateCompletionEditableProjectIdentity(completion), /request identity conflicts/);
});

for (const contentType of ["QI", "FPI"]) {
  test(`${contentType} approval stays one record with all cards and project identity`, () => {
    const { buildPoetryPleaseGraphicRecord } = loadFunctions(["buildPoetryPleaseGraphicRecord"]);
    const result = buildPoetryPleaseGraphicRecord({
      graphicsQcDecision: "APPROVE", contentType, assetLinkUrl: assets[0].assetUrl,
      assetPreviewUrl: assets[0].assetPreviewUrl, assetSet: assets,
      sourceRecordId: "synthetic-source", graphicsRequestId: "synthetic-request",
      pigCompletionId: "synthetic-completion", pigProjectId: "synthetic-project",
      editableProjectFileId: "synthetic-json", editableProjectUrl: "https://example.com/project.json"
    });
    assert.deepEqual(result.assetSet, assets);
    assert.equal(result.contentType, contentType);
    assert.equal(result.sourceRecordId, "synthetic-source");
    assert.equal(result.sourceRequestId, "synthetic-request");
    assert.equal(result.sourceCompletionId, "synthetic-completion");
    assert.equal(result.pigProjectId, "synthetic-project");
    assert.equal(result.editableProjectFileId, "synthetic-json");
    assert.equal(result.driveLink, assets[0].assetUrl);
  });
}

test("synthetic canary stays marked through QC overlay and Poetry Please export", () => {
  const context = loadFunctions(["overlayRuntimeHandoffState", "buildPoetryPleaseGraphicRecord"], {
    normalizePoetryPleaseContentType: (value, fallback) => value || fallback
  });
  const record = context.overlayRuntimeHandoffState({
    graphicsRequestId: "canary-request", contentType: "QI", assetLinkUrl: assets[0].assetUrl,
    sourceRecordId: "canary-source", pigCompletionId: "canary-completion"
  }, {
    "canary-request": {
      qcStatus: "approved", assetUrl: assets[0].assetUrl, assetSet: assets,
      pigPayload: { syntheticTest: true, assetSet: assets }
    }
  });
  assert.equal(record.syntheticTest, true);
  const exported = context.buildPoetryPleaseGraphicRecord(record);
  assert.equal(exported.syntheticTest, true);
  assert.deepEqual(exported.assetSet, assets);
});

test("QC approval restores the saved multi-card completion before Poetry Please export", async () => {
  const context = loadFunctions(["hydrateApprovedFirestoreGraphicRecord", "buildPoetryPleaseGraphicRecord"], {
    syncWeaverRuntimeDb: async (action, query) => {
      assert.equal(action, "get_graphics_completion");
      assert.equal(query.completionId, "canary-completion");
      return { record: {
        id: "canary-completion", graphicsRequestId: "canary-request",
        sourcePayload: {
          contentType: "QI", assetUrl: assets[0].assetUrl,
          assetPreviewUrl: assets[0].assetPreviewUrl, assetSet: assets,
          sourceRecordId: "canary-source", pigProjectId: "canary-project",
          editableProjectFileId: "canary-json", syntheticTest: true
        }
      } };
    }
  });
  const approved = await context.hydrateApprovedFirestoreGraphicRecord({
    graphicsRequestId: "canary-request", pigCompletionId: "canary-completion",
    graphicsQcDecision: "APPROVE", assetLinkUrl: assets[0].assetUrl
  });
  const exported = context.buildPoetryPleaseGraphicRecord(approved);
  assert.deepEqual(exported.assetSet, assets);
  assert.equal(exported.syntheticTest, true);
  assert.equal(exported.sourceRecordId, "canary-source");
  assert.equal(exported.sourceRequestId, "canary-request");
  assert.equal(exported.sourceCompletionId, "canary-completion");
  assert.equal(exported.editableProjectFileId, "canary-json");
});

test("FPI outbound import retains all cards and editable identity after its payload mapping", async () => {
  let posted;
  const context = loadFunctions(["handoffApprovedGraphicsToPoetryPlease"], {
    fetch: async (_url, options) => {
      posted = JSON.parse(options.body);
      return { ok: true, status: 200, json: async () => ({ ok: true }) };
    }
  });
  await context.handoffApprovedGraphicsToPoetryPlease([{
    contentType: "FPI", sourceRecordId: "synthetic-source", assetSet: assets,
    sourceRequestId: "original-request", sourceCompletionId: "original-completion",
    pigProjectId: "whole-project", editableProjectFileId: "whole-project-json",
    driveLink: assets[0].assetUrl
  }]);
  assert.equal(posted.records.length, 1);
  assert.deepEqual(posted.records[0].assetSet, assets);
  assert.equal(posted.records[0].editableProjectFileId, "whole-project-json");
  assert.equal(posted.records[0].sourceRequestId, "original-request");
});

test("legacy single-card FPI keeps its existing generated source ID", () => {
  const { buildPoetryPleaseGraphicRecord } = loadFunctions(["buildPoetryPleaseGraphicRecord"]);
  const record = buildPoetryPleaseGraphicRecord({
    graphicsQcDecision: "APPROVE", contentType: "FPI", assetLinkUrl: assets[0].assetUrl,
    sourceRecordId: "legacy-sheet-row", bookShortener: "TEST", poemTitle: "Poem"
  });
  assert.equal(record.sourceRecordId, "TEST-FPI-Poem");
  assert.equal(record.driveLink, assets[0].assetUrl);
  assert.equal(record.assetSet, undefined);
});

test("unexpected FP graphic is rejected instead of being sent as QI", async () => {
  const context = loadFunctions(["buildPoetryPleaseGraphicRecord", "handoffApprovedGraphicsToPoetryPlease"]);
  assert.throws(() => context.buildPoetryPleaseGraphicRecord({
    graphicsQcDecision: "APPROVE", contentType: "FP", assetLinkUrl: assets[0].assetUrl, assetSet: assets
  }), /Unsupported approved graphic contentType: FP/);
  await assert.rejects(context.handoffApprovedGraphicsToPoetryPlease([{ contentType: "FP", assetSet: assets }]),
    /Unsupported approved graphic contentType: FP/);
});

test("repair return posts ordered replacements and affected card through the actual status serializer", async () => {
  let posted;
  const context = loadFunctions(["buildStructuredRepairReturn", "mergeSingleCardRepairReturn", "updatePoetryPleaseRepairStatus"], {
    fetch: async (_url, options) => {
      posted = JSON.parse(options.body);
      return { ok: true, status: 200, json: async () => ({ ok: true }) };
    }
  });
  const original = assets.map(card => ({ ...card, assetPreviewUrl: `https://pp.example.com/${card.index}` }));
  const replacement = { ...assets[1], assetFileId: "new-card-2", assetUrl: "https://example.com/new-card-2.png" };
  const structured = context.mergeSingleCardRepairReturn(context.buildStructuredRepairReturn({
    assetSet: [assets[0], replacement, assets[2]], affectedCardIndex: 2, contentType: "FPI",
    sourceRecordId: "synthetic-source", sourceRequestId: "original-request",
    sourceCompletionId: "original-completion", pigProjectId: "synthetic-project",
    replacementAssetId: assets[0].assetFileId, replacementAssetLink: assets[0].assetUrl,
    pigCompletionId: "replacement-completion", weaverJobId: "repair-job"
  }), original);
  await context.updatePoetryPleaseRepairStatus("synthetic-repair", "returned", "", structured);
  assert.deepEqual(posted.assetSet, [original[0], replacement, original[2]]);
  assert.equal(posted.affectedCardIndex, 2);
  assert.equal(posted.contentType, "FPI");
  assert.equal(posted.sourceCompletionId, "original-completion");
  assert.equal(posted.pigJobId, "replacement-completion");
  assert.equal(posted.replacementAssetId, replacement.assetFileId);
  assert.equal(posted.replacementAssetLink, replacement.assetUrl);
  assert.throws(() => context.mergeSingleCardRepairReturn(
    { ...structured, assetSet: [replacement, replacement, assets[2]] }, original
  ), /unaffected_card_file_changed/);
});
