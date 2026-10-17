import test from "node:test";
import assert from "node:assert/strict";
import { fingerprintExcerptForCatalogKey, poemKeyForExcerptText } from "../poem_key_seed.mjs";

test("catalog seed uses Weaver's normalized excerpt hash", () => {
  assert.equal(
    fingerprintExcerptForCatalogKey("“Don’t *go*—now.”"),
    "64263adaf05474981442389a723b34bae61eef8e61843d59caaee8575484e965"
  );
  assert.equal(
    poemKeyForExcerptText("“I know a quiet, but don’t know who or what you were, I am, was.”"),
    "BPP-04467"
  );
  assert.equal(poemKeyForExcerptText(""), null);
  assert.equal(poemKeyForExcerptText("A phrase absent from the verified seed"), null);
});
