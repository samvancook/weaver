# Catalog Identity and Reconciliation Roadmap

## Established

- `unreliable narrator` / `UN` is the canonical identity. Do not recreate `all the ugly bits` / `ATUB` as a separate book.
- Weaver's hosted intake reads this repository's `data/formal_catalog.db`. The local poetry-catalog database and the workbench branch are inputs for reconciliation, not hosted truth.
- The hosted catalog currently keeps `unreliable narrator` as `set_aside`; do not make it selectable merely by importing poem text.
- A read-only scan of 381 active Weaver graphics-handoff records found no `ATUB` or `all the ugly bits` matches.
- The repository catalog database also contains no old-name or old-shortener match.
- No Poetry Please record was deleted, renamed, or otherwise changed by this work. The signed-in production Content Library search on 2026-09-24 returned 137 records for `unreliable narrator`, all with that book: 22 EXC, 50 FP, 11 FPI, and 54 QI. Searches for `all the ugly bits` and `ATUB` each returned no matching content.
- Poetry Please's canonical repository lookup has one `unreliable narrator` / `UN` book entry with `all the ugly bits` as a `titleKeys` alias. This is code-only metadata, not evidence of live content. Keep the alias until production records are audited, then decide whether it can be retired.

## Next

1. Done for the live Poetry Please Content Library view: use the current signed-in Chrome admin tab to search both names. The new tab was accessible; 137 canonical-name results and zero old-title/shortener results were found. No in-place rename is indicated by this view. Keep the code-only old-title alias under review; retire it only after confirming no other production lookup depends on it. The Content Library search is not a raw Firestore scan, so do not generalize the zero-match result to every possible internal collection.
2. Reconcile the nine remaining workbench-only source editions against existing canonical books and publication decisions. Eight belong to books marked `set_aside`. `an everyday occurrence` is `catalog_ok`, but Weaver already has 52 rows from a PDF source; the additional EPUB has 25 rows, 16 with exact title matches. Do not import that EPUB as a new book or bulk-replace the hosted database. Compare editions and import only verified missing poems, preserving status and source provenance.
3. Done in the repository catalog and deployed in Weaver revision `weaver-00436-scc`: reconcile `unreliable narrator` to 46 distinct poem sections. The source DOCX has 48 numbered TOC entries; two do not have distinct body poems. Split `cycle` and `namesake` from merged rows, remove the duplicated vase-poem text from `universal truths`, and exclude the prose `prologue`. The `set_aside` status and intake visibility remain unchanged. The local endpoint returns 46; external hosted read verification remains pending because the shell cannot resolve the service URL and the browser client blocked it.
4. Produce a per-book report of imported, matched, excluded non-poem sections, unresolved, and status-preserved rows. Verify hosted and local catalog lookups use the intended repository database after deployment.

## Remaining Workbench Sources

No source edition below has been imported from the workbench. Their extracted counts match workbench metadata, but source quality and overlap have not yet been fully validated.

| Book | Extracted | Hosted status |
| --- | ---: | --- |
| A Choir of Honest Killers | 60 | set_aside |
| DON'T BE AFRAID TO BE BAD | 67 | set_aside |
| Flee | 20 | set_aside |
| Living at Baggage Claim | 48 | set_aside |
| Roads | 63 | set_aside |
| Stunt Water | 63 | set_aside |
| Tooth Gaps in the Archives | 23 | set_aside |
| an everyday occurrence (additional EPUB) | 25 | catalog_ok; 52 PDF rows already present, 16 exact-title overlaps |
| without the frills | 53 | set_aside |

## Completion Checks

- No active Weaver or Poetry Please record uses the old book identity.
- No duplicate UN book identity was created; original content/request IDs remain stable.
- Hosted poem/title results agree with validated source sections while publication statuses remain intentional.
- Local diagnostic scripts and hosted intake have documented, explicit database paths.
