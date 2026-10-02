import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { runInNewContext } from "node:vm";

const source = readFileSync(new URL("../public/app.js", import.meta.url), "utf8");
const start = source.indexOf("function getPendingRecordsForBookKey(");
const end = source.indexOf("function isVideoReviewRecord(", start);
assert.ok(start >= 0 && end > start);
const { getPendingRecordsForBookKey } = runInNewContext(
  `${source.slice(start, end)}\n({ getPendingRecordsForBookKey })`,
  {
    REVIEW_VIDEO_EVENT_KEY: "__video_event__:",
    REVIEW_VIDEOS_BOOK_KEY: "__video_excerpts__",
    getPendingVideoRecords: records => records.filter(record => record.intakeMode === "video"),
    extractGoogleDriveFileId: () => "",
    normalizeBookKey: value => value.toLowerCase()
  }
);

test("event-wide video review includes every pending video excerpt in that event", () => {
  const records = [
    { intakeMode: "video", sourceEvent: "Charm City", sourceVideoFileId: "video-1", recordId: "pull-1" },
    { intakeMode: "video", sourceEvent: "Charm City", sourceVideoFileId: "video-2", recordId: "pull-2" },
    { intakeMode: "video", sourceEvent: "Other", sourceVideoFileId: "video-3", recordId: "pull-3" },
    { intakeMode: "book", sourceEvent: "Charm City", recordId: "book-1" }
  ];
  assert.deepEqual(
    Array.from(getPendingRecordsForBookKey("__video_event__:Charm%20City", records), record => record.recordId),
    ["pull-1", "pull-2"]
  );
  assert.deepEqual(
    Array.from(getPendingRecordsForBookKey("__video_excerpts__:video-1", records), record => record.recordId),
    ["pull-1"]
  );
});
