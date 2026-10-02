import test from "node:test";
import assert from "node:assert/strict";
import { findApprovedTextExcerptMatch } from "../video_excerpt_folding.mjs";

const video = {
  author: "Poet", bookTitle: "Book", poemTitle: "Poem", excerptText: "First line\nSecond line"
};
const approved = {
  ...video, sourceRecordId: "weaver:book-1", sourceVideoFileId: null,
  approval: { approvedForUse: true }
};

test("folds one approved text excerpt with identical identity and text", () => {
  assert.equal(findApprovedTextExcerptMatch(video, [approved])?.sourceRecordId, "weaver:book-1");
  assert.equal(findApprovedTextExcerptMatch({ ...video, excerptText: "First line Second line" }, [approved])?.sourceRecordId, "weaver:book-1");
});

test("leaves unapproved, ambiguous, and different-identity matches for review", () => {
  assert.equal(findApprovedTextExcerptMatch(video, [{ ...approved, approval: { approvedForUse: false } }]), null);
  assert.equal(findApprovedTextExcerptMatch(video, [approved, { ...approved, sourceRecordId: "weaver:book-2" }]), null);
  assert.equal(findApprovedTextExcerptMatch({ ...video, poemTitle: "Other poem" }, [approved]), null);
  assert.equal(findApprovedTextExcerptMatch({ ...video, bookTitle: "" }, [approved]), null);
  assert.equal(findApprovedTextExcerptMatch(video, [{ ...approved, sourceVideoFileId: "video-1" }]), null);
});
