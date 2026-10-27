import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";

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
  const context = loadFunctions(["buildStructuredRepairReturn", "updatePoetryPleaseRepairStatus"], {
    fetch: async (_url, options) => {
      posted = JSON.parse(options.body);
      return { ok: true, status: 200, json: async () => ({ ok: true }) };
    }
  });
  const structured = context.buildStructuredRepairReturn({
    assetSet: assets, affectedCardIndex: 2, contentType: "FPI",
    sourceRecordId: "synthetic-source", sourceRequestId: "original-request",
    sourceCompletionId: "original-completion", pigProjectId: "synthetic-project",
    replacementAssetId: assets[0].assetFileId, replacementAssetLink: assets[0].assetUrl,
    pigCompletionId: "replacement-completion", weaverJobId: "repair-job"
  });
  await context.updatePoetryPleaseRepairStatus("synthetic-repair", "returned", "", structured);
  assert.deepEqual(posted.assetSet, assets);
  assert.equal(posted.affectedCardIndex, 2);
  assert.equal(posted.contentType, "FPI");
  assert.equal(posted.sourceCompletionId, "original-completion");
  assert.equal(posted.pigJobId, "replacement-completion");
  assert.equal(posted.replacementAssetLink, assets[0].assetUrl);
});
