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
