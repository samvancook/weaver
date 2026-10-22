# Catalog Identity and Reconciliation Roadmap

## Established

- Weaver's book intake has two separate inputs. The publishing-order Sheet provides book metadata and suggestions; catalog-backed poem titles and validation come from `data/formal_catalog.db`. A book appearing in suggestions does not prove its poems are available.
- `unreliable narrator` / `UN` is the canonical identity. Do not recreate `all the ugly bits` / `ATUB` as a separate book.
- Weaver's hosted intake reads this repository's `data/formal_catalog.db`. The local poetry-catalog database and the workbench branch are inputs for reconciliation, not hosted truth.
- The hosted catalog currently keeps `unreliable narrator` as `set_aside`; do not make it selectable merely by importing poem text.
- A read-only scan of 381 active Weaver graphics-handoff records found no `ATUB` or `all the ugly bits` matches.
- The repository catalog database also contains no old-name or old-shortener match.
- No Poetry Please record was deleted, renamed, or otherwise changed by this work. The signed-in production Content Library search on 2026-09-24 returned 137 records for `unreliable narrator`, all with that book: 22 EXC, 50 FP, 11 FPI, and 54 QI. Searches for `all the ugly bits` and `ATUB` each returned no matching content.
- Poetry Please's canonical repository lookup has one `unreliable narrator` / `UN` book entry with `all the ugly bits` as a `titleKeys` alias. The alias is an input compatibility key: old-title imports resolve to canonical UN metadata and product links. Keep it until historical/external inputs are audited; it is not a separate live book or public display title.

## New-book intake path and BBFR gap (2026-10-22)

Previously deployed path: identify the authoritative EPUB or DOCX and its edition; extract and reconcile book identity, poem titles, text, and source provenance into the repository's tracked `data/formal_catalog.db`; deploy a new Weaver image; then read back `/api/intake/options` and `/api/intake/catalog/poems?bookTitle=...` from the live service. `catalog_validate.py` and `intake_catalog_lookup.py` read that bundled SQLite file. `.gcloudignore`, `.dockerignore`, and `Dockerfile` include it in the deployment image. The publishing-order Sheet supplies metadata, not poem text. Firestore's Weaver operational ledger is a separate store. A runtime overlay that imports selected approved books from the main Catalog is under implementation to remove new manuscript text from Git history; do not describe it as live until deployed and read back.

Observed BBFR state on 2026-10-22: the live publishing-order suggestions include `BLACK BOX FLIGHT RECORDER` by Aishatu Ado (`BBFR`, Spring 2027), but the live catalog-backed book list excludes it and its poem endpoint returns `Book not found in catalog.` The main [Button Poetry Catalog service](https://button-poetry-catalog-350789123099.us-central1.run.app/books/BBFR) now identifies a `catalog_ok` DOCX edition, `BBFR - Ado - 10142026 - Final Proofing.docx`, with 26 poems. A validated runtime-overlay import of those 26 poems is prepared in this repository checkout; local SQLite readback matches all 26 source titles and exact served texts. **Live Weaver remains unresolved until deployment and live readback.** `WG` / `Worm Girl` is present in publishing-order suggestions but the main Catalog currently returns `Book not found` for both names, so it cannot be imported yet.

The main Catalog response includes `poem_key` values for BBFR, but Weaver's current `catalog_poems` table has no key column and the bundled 2026-10-15 poem-key seed does not include BBFR. The book/poem lookup can work after deployment, while stable poem-key propagation remains a separate schema and seed-refresh task. The public Weaver GitHub repository currently tracks `data/formal_catalog.db`; publishing a refreshed copy there would expose manuscript-derived poem text. The new overlay path downloads approved books into a private ephemeral database at Weaver startup without committing new poem texts to Git. Verify this runtime dependency and its restart behavior before calling the live path complete.

Make the next import repeatable:

1. Record the source file link, edition/version, author/title/shortener, publication status, and permission to use the text before extraction. Locate the authoritative catalog source rather than inferring it from the publishing-order row.
2. Use one documented import command or script to build or update Weaver's runtime catalog from that source. The repository historically had one-off reconciliation scripts and a bundled database; the new overlay is a first targeted import path, not yet a general intake service.
3. Produce a review report showing source sections against imported poem count and titles, duplicate titles, empty or implausible text, and excluded non-poem sections. Require an explicit status decision: an intentionally `set_aside` edition must not appear as selectable merely because text was extracted.
4. Check `book_status`, `catalog_books`, and `catalog_poems` locally, including exact author/title/shortener and sample poem text. Then deploy the changed catalog artifact and verify the actual live revision, catalog book option, poem count, and representative titles/text through live endpoints. Record the source ID, catalog artifact hash, deployment revision, and readback in the handoff.
5. Show a distinct `metadata only / no catalog poems` state in Weaver when a publishing-order book has no catalog identity, so selecting BBFR does not lead to a generic `Book not found` surprise. Add a freshness or coverage check that compares upcoming publishing-order books with packaged catalog identities and flags missing sources before excerpt work begins.

Open decisions: whether catalog import should remain a repository build artifact or move to a managed shared store; who reviews extraction and activates each edition. The main Catalog service is live, but Weaver's own poem lookup still uses its bundled SQLite snapshot until code and deployment change that architecture.

## Next

### Shared catalog poem keys

- Small slice implemented locally: backfill `excerpt_entries.poem_key` from the catalog's 2026-10-15 seed by verified unique excerpt hash; ship a compact, versioned hash-to-key snapshot so approved Weaver excerpt handoffs can include `poemKey` and `poemKeySource` without guessing from a title. Unmatched and ambiguous excerpts keep a null key. Verify a live handoff and Poetry Please readback after deployment.
- Firestore `excerptRecords` backfill executed 2026-10-16 against `button-weaver-internal/weaverledger`: 848 records scanned; 10 matched a verified catalog key and were conditionally patched with only `poemKey` and `poemKeySource`, then individually read back. Final dry run: 0 planned, 10 already keyed, 4 missing excerpt text, 834 without a verified key. No conflicting key or failed write was observed. Reusable dry-run-first command: `backfill_firestore_excerpt_poem_keys.py` using the operational excerpt database path; use the documented named Weaver deployer credential, not user ADC.
- Larger reconciliation still needed for the 834: 803 are video-linked and 660 have no book title, so many cannot be tied to a catalog poem from the current seed. Confirm whether Poetry Please exposes an exact `sourceRecordId` to `poemKey` mapping for these records, audit conflicts against Weaver text/source identity, and backfill only unambiguous keys. Also audit existing handoff documents, confirm Poetry Please's importer persists newly supplied keys, add a refresh path for later catalog seed releases, a live catalog lookup for new excerpts, and video-to-poem key propagation. Keep catalog row IDs out of identity joins and do not infer keys from titles.

Open provenance issue: reconcile Weaver and Poetry Please release-catalog labels. Identify the authoritative catalog/edition field for each system and compare a small set of exact book identities before changing labels. Do not overwrite book metadata or infer a catalog from a repair request, queue position, or display text.

1. Done for the live Poetry Please Content Library view: use the current signed-in Chrome admin tab to search both names. The new tab was accessible; 137 canonical-name results and zero old-title/shortener results were found. No in-place rename is indicated by this view. Keep the code-only old-title compatibility alias until source inputs and external handoffs no longer send the old title. The Content Library search is not a raw Firestore scan, so do not generalize the zero-match result to every possible internal collection.
2. Decision on the eight `set_aside` editions: no import now. Each book is present in `canonical_books`, has `effective_status=set_aside`, and intentionally has zero hosted `catalog_books` and `catalog_poems` rows. Their workbench extracts are source material, not missing active catalog records. Revisit only if a book's publication/intake status changes.
3. One active source-quality question remains: `an everyday occurrence` has 52 hosted PDF rows and an additional 25-row EPUB. The EPUB comparison found 16 exact-title matches, two punctuation-only title matches (`Don't RSVP to My Pity Party`, `You Can't Convince Me`), two poems whose text appears inside other PDF-extraction rows (`I Talk to the Moon Instead of the Police`, `Cause of Death`), and five without matching first/last 12-word anchors (`Dedication`, `The Surrogate Mother Speaks About Her Coming of Age Story`, `Inheritance`, `Confucius Was Not a Feminist`, `Dragon Lady Finds Tinder`). The PDF extraction also contains malformed title boundaries, so a row-count increase alone would not improve accuracy. Compare these five to the intended published edition and repair boundaries before any selective import; preserve status and source provenance.
4. Done in merged Weaver PR #12 and deployed in revision `weaver-00436-scc`: reconcile `unreliable narrator` to 46 distinct poem sections. The source DOCX has 48 numbered TOC entries; two do not have distinct body poems. Split `cycle` and `namesake` from merged rows, remove the duplicated vase-poem text from `universal truths`, and exclude the prose `prologue`. A direct hosted read verified 46 titles, including `cycle` and `namesake`; hosted intake options still exclude UN. The `set_aside` status remains unchanged.

## Remaining Workbench Sources

No source edition below has been imported from the workbench. The eight `set_aside` editions require no active-catalog action; only the AEO edition comparison remains open.

| Book | Extracted | Hosted status |
| --- | ---: | --- |
| A Choir of Honest Killers | 60 | set_aside |
| DON'T BE AFRAID TO BE BAD | 67 | set_aside |
| Flee | 20 | set_aside |
| Living at Baggage Claim | 48 | set_aside |
| Roads | 63 | set_aside |
| Stunt Water | 63 | set_aside |
| Tooth Gaps in the Archives | 23 | set_aside |
| an everyday occurrence (additional EPUB) | 25 | catalog_ok; 52 PDF rows already present, 16 exact-title overlaps, 2 punctuation-only matches, 2 text overlaps, 5 unresolved |
| without the frills | 53 | set_aside |

## Completion Checks

- No active Weaver or Poetry Please record uses the old book identity.
- No duplicate UN book identity was created; original content/request IDs remain stable.
- Hosted poem/title results agree with validated source sections while publication statuses remain intentional.
- Local diagnostic scripts and hosted intake have documented, explicit database paths.
