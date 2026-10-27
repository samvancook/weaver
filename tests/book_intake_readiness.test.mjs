import assert from "node:assert/strict";
import test from "node:test";
import { buildPublishingBookReadiness } from "../book_intake_readiness.mjs";

const publishing = [
  { title: "Worm Girl", bookShortener: "WG", releaseCatalog: "Spring 2027" },
  { title: "New Book", bookShortener: "NB", releaseCatalog: "Spring 2027" }
];

test("live catalog coverage distinguishes readable books from metadata only", () => {
  assert.deepEqual(buildPublishingBookReadiness(publishing,
    [{ title: "Worm Girl", bookShortener: "WG" }], "live").map(book => book.status),
  ["catalog_backed", "metadata_only"]);
});

test("a Catalog outage leaves coverage unverified instead of falsely absent", () => {
  assert.deepEqual(buildPublishingBookReadiness(publishing, [], "cached_fallback")
    .map(book => book.status), ["unverified", "unverified"]);
});
