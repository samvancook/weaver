#!/usr/bin/env node
// Read-only production check. Never submits an excerpt or changes a queue.
const base = (process.env.WEAVER_URL || "https://weaver.buttonpoetry.com").replace(/\/$/, "");
const book = process.env.WEAVER_CHECK_BOOK || "Roads";
const releaseCatalog = process.env.WEAVER_CHECK_RELEASE_CATALOG || "Spring 2027";
const timeoutMs = 20000;

async function request(path, options = {}) {
  const response = await fetch(new URL(path, base), {
    ...options,
    signal: AbortSignal.timeout(timeoutMs),
  });
  if (!response.ok) throw new Error(`${path}: HTTP ${response.status}`);
  return response;
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

try {
  const options = await (await request("/api/intake/options")).json();
  assert(options.ok && Array.isArray(options.catalogBooks), "Catalog book options unavailable");
  assert(options.catalogBooks.some(row => row.title === book), `${book} is absent from Catalog-backed options`);
  const releaseReadiness = (options.publishingBookReadiness || [])
    .filter(row => row.releaseCatalog === releaseCatalog)
    .map(row => ({ title: row.title, bookShortener: row.bookShortener, status: row.status }));

  const poemPath = `/api/intake/catalog/poems?bookTitle=${encodeURIComponent(book)}`;
  const poems = await (await request(poemPath)).json();
  assert(poems.ok && Array.isArray(poems.poemOptions) && poems.poemOptions.length,
    `${book} has no selectable poems`);
  const selected = poems.poemOptions.find(row => row.poemKey && row.title) || poems.poemOptions[0];
  const previewPath = `/catalog-poem?bookTitle=${encodeURIComponent(book)}`
    + `&poemTitle=${encodeURIComponent(selected.title)}`
    + `&poemKey=${encodeURIComponent(selected.poemKey || "")}`;
  const preview = await (await request(previewPath)).text();
  assert(preview.includes(selected.title) && !preview.includes("Poem View Unavailable"),
    `${book}: selected poem did not render`);

  const pending = await (await request("/api/review/pending-records")).json();
  assert(pending && typeof pending === "object" && pending.ok !== false,
    "Review queue read failed");
  await request("/graphics-handoff/queue");

  console.log(JSON.stringify({ok: true, base, book, catalogBooks: options.catalogBooks.length,
    releaseCatalog, releaseReadiness,
    poemOptions: poems.poemOptions.length, previewPoemKey: selected.poemKey || null,
    reviewQueue: "readable", graphicsQueue: "readable", sheetSubmission: "not tested"}, null, 2));
} catch (error) {
  console.error(JSON.stringify({ok: false, base, book, error: error.message}));
  process.exitCode = 1;
}
