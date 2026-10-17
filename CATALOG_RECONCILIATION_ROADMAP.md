# Catalog Identity and Reconciliation Roadmap

## Established

- `unreliable narrator` / `UN` is the canonical identity. Do not recreate `all the ugly bits` / `ATUB` as a separate book.
- Weaver's hosted intake reads this repository's `data/formal_catalog.db`. The local poetry-catalog database and the workbench branch are inputs for reconciliation, not hosted truth.
- The hosted catalog currently keeps `unreliable narrator` as `set_aside`; do not make it selectable merely by importing poem text.
- A read-only scan of 381 active Weaver graphics-handoff records found no `ATUB` or `all the ugly bits` matches.
- The repository catalog database also contains no old-name or old-shortener match.
- No Poetry Please record was deleted, renamed, or otherwise changed by this work. The signed-in production Content Library search on 2026-09-24 returned 137 records for `unreliable narrator`, all with that book: 22 EXC, 50 FP, 11 FPI, and 54 QI. Searches for `all the ugly bits` and `ATUB` each returned no matching content.
- Poetry Please's canonical repository lookup has one `unreliable narrator` / `UN` book entry with `all the ugly bits` as a `titleKeys` alias. The alias is an input compatibility key: old-title imports resolve to canonical UN metadata and product links. Keep it until historical/external inputs are audited; it is not a separate live book or public display title.

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
