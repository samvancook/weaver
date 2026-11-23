#!/usr/bin/env node
// Read-only production check. Never submits an excerpt or changes a queue.
const base = (process.env.WEAVER_URL || "https://weaver.buttonpoetry.com").replace(/\/$/, "");
const book = process.env.WEAVER_CHECK_BOOK || "Roads";
const releaseCatalog = process.env.WEAVER_CHECK_RELEASE_CATALOG || "Spring 2027";
const seasonMode = process.argv.includes("--season");
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
  const releaseReadiness = (options.publishingBookReadiness || [])
    .filter(row => row.releaseCatalog === releaseCatalog)
    .map(row => ({ title: row.title, bookShortener: row.bookShortener, status: row.status }));

  const targets = seasonMode ? releaseReadiness : [{title:book,status:"check"}];
  assert(targets.length, `${releaseCatalog} has no publishing-book entries`);
  const results = [];
  for (const target of targets) {
   const book = target.title;
   try {
   assert(options.catalogBooks.some(row => row.title === book), `${book} is absent from Catalog-backed options (${target.status})`);
   const poemPath = `/api/intake/catalog/poems?bookTitle=${encodeURIComponent(book)}`;
  const poems = await (await request(poemPath)).json();
  assert(poems.ok && Array.isArray(poems.poemOptions) && poems.poemOptions.length,
    `${book} has no selectable poems`);
  assert(poems.poemOptions.every(row => row.poemKey && row.title), `${book} has unkeyed poem choices`);
  const keys = poems.poemOptions.map(row => row.poemKey);
  assert(new Set(keys).size === keys.length, `${book} repeats a poem key; duplicate titles must retain distinct identities`);
  const selected = poems.poemOptions[0];
  const previewPath = `/catalog-poem?bookTitle=${encodeURIComponent(book)}`
    + `&poemTitle=${encodeURIComponent(selected.title)}`
    + `&poemKey=${encodeURIComponent(selected.poemKey || "")}`;
  const preview = await (await request(previewPath)).text();
  assert(preview.includes(selected.title) && !preview.includes("Poem View Unavailable"),
    `${book}: selected poem did not render`);
   results.push({book,ok:true,poemOptions:poems.poemOptions.length,previewPoemKey:selected.poemKey});
   } catch (error) { results.push({book,ok:false,status:target.status,error:error.message}); }
  }

  const pending = await (await request("/api/review/pending-records")).json();
  assert(pending && typeof pending === "object" && pending.ok !== false,
    "Review queue read failed");
  await request("/graphics-handoff/queue");

  const ok = results.every(row => row.ok);
  console.log(JSON.stringify({ok, base, book:seasonMode?null:book, catalogBooks: options.catalogBooks.length,
    releaseCatalog, releaseReadiness, results,
    ...(!seasonMode&&results[0]?.ok?{poemOptions:results[0].poemOptions,previewPoemKey:results[0].previewPoemKey}:{}),
    reviewQueue: "readable", graphicsQueue: "readable", sheetSubmission: "not tested",
    scope:"Book choices and one keyed poem preview per Catalog-backed book; no import, approval, excerpt submission or publication"}, null, 2));
  if (!ok) process.exitCode = 1;
} catch (error) {
  console.error(JSON.stringify({ok: false, base, book, error: error.message}));
  process.exitCode = 1;
}
