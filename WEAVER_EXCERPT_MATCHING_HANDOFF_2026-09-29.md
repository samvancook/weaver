# Weaver excerpt matching handoff (2026-09-29)

## Decision and boundary

Excerpt Database work in this thread is at a read-only verification checkpoint. Weaver owns any change to its live duplicate-checking or approved-excerpt consumer. Do not treat the Google Sheet workflow bridge, Weaver's Firestore `excerptRecords`, and the canonical matching export as the same store.

## Verified artifacts

- Local operational Excerpt DB: `data/excerpt_library.db`, 47,461 source rows in the verified snapshot; file last modified 2026-09-23.
- Local normalized DB: `data/excerpt_library_normalized.db`, 32,666 canonical excerpts.
- Matching export generated 2026-09-28: `data/exports/weaver_excerpt_library/manifest.json` (schema v4), `weaver_excerpt_library.json`, `weaver_excerpt_library.csv`, and `operational_excerpt_index.csv`.
- Shared Drive delivery folder: https://drive.google.com/drive/folders/11i6wGukLM6MtzsCRu3k1w55fGSDKC2Vg
- Drive manifest: https://drive.google.com/file/d/1D9fb78UBSzpYLI6X_cEzDH2C_NT0oHT2/view
- Canonical JSON SHA-256: `95c503df0f150b5b0971a30daf17523e7b315f03e1a35d9cb7e80f6c507e7a59`.
- Canonical CSV SHA-256: `c5770082b843a4e42f72cfb1d2a9c3eb63c7543e7a892bbb1e4b935789add15d`.
- Operational index CSV SHA-256: `33db4c6bc5cd8872284dc9e721dbcb60f24beb09c9de1efd6e952e8b22b847aa`.
- The local files match the manifest hashes. The Drive manifest contains the same values and uploaded file sizes match locally; Drive file contents were not independently hashed.

The canonical export is one row per normalized excerpt for matching. The operational index retains all 47,461 source rows for coverage and provenance. The manifest accounts for 634 source rows outside the canonical feed (394 COV, 103 INT_LAYOUT, 137 likely non-excerpts), with zero unaccounted rows. This is a local snapshot, not proof of a live consumer refresh.

## Separate production bridge

The Button Spreadsheet Toolkit Apps Script matcher uses the 48-column workflow bridge, not the canonical matching CSV:

https://docs.google.com/spreadsheets/d/1fARFnzlHt6ekajeubmkw4DIlZnQ4Pv_LIuKbdouGDLY/edit

The bridge's 2026-09-23 publish manifest records 47,461 data rows and CSV SHA-256 `155de6c629c9512f91e2c125135d681864a9a9c7401adf19f4a6f6039cf98513`. Regenerating that compatibility CSV from the current local DB with the historical named exporter produced the identical hash. The live sheet has the expected 48-column header, last data row at 47,462, and no values in rows 47,463-47,464. The newer 2026-09-28 matching export must not be pasted into this bridge merely because its file date is newer.

## Approved-feed checkpoint

The named historical `sync_weaver_approved_feed.py` at git commit `4764253` was run in dry-run mode only. Against `https://weaver.buttonpoetry.com/api/excerpts/approved/export` with saved cursor `2026-07-02T19:34:04.000Z`, it returned one boundary record, `weaver:row-10852`, with the same `sourceUpdatedAt`; that source ID is already in the local DB. This shows no newer record in that feed response. It does not prove that every possible Weaver approval is in the feed. No DB, sync cursor, bridge, or Weaver data was changed.

The named sync, workflow exporter, and bridge publisher exist in git history at `4764253` but are absent from the current checkout. Do not deploy or restore historical Weaver work without reconciling it with current code and the approved workflow.

## Weaver action requested

### Read-only code trace (2026-09-29)

The current repository does have an existing library-match path. `catalog_validate.py` populates `libraryExcerptMatch` by calling `find_library_excerpt_match` from `excerpt_library.py` when its local SQLite connection is available; `public/app.js` renders the resulting match in review. This path does not reference the September 28 Drive JSON/CSV or manifest. `Dockerfile` copies `data`, while `.gcloudignore` includes `data/excerpt_library_runtime.db` but does not include the full `data/excerpt_library.db`. Thus the deployment package does not establish that the full local library is available to that validator. This is repository/package evidence only; the actual deployed revision and runtime lookup result still need a Weaver-owned check.

The pinned Codex thread titled `Weaver` is currently working on a separate video-handoff change. Do not interrupt or conflate that work with this matching-export handoff.

1. Identify the live Weaver duplicate-check consumer and its current data source. Do not infer usage from an export existing in Drive.
2. If it does not load this matching export, make the smallest Weaver-owned change through the approved deployment workflow. Match on `normalizedExcerpt`; use `operational_excerpt_index.csv` when source-row provenance or complete coverage is required. Do not equate graphics eligibility with a completed asset.
3. Produce a consumer receipt: loaded artifact name and SHA-256, manifest schema/version or generated time, load time, loaded row count, and a successful lookup test (for example Rudy Francisco's "Tragedy and silence have the exact same address"). If the consumer cannot expose a receipt, report that capability gap rather than claiming the export is live.
4. Keep legacy-sheet cutover separate. Do not freeze the old sheet solely on the basis of this matching-export verification.

## Excerpt Database next step

After Weaver returns a receipt, reconcile the historical approved-feed sync and publication scripts with the current checkout, then establish a repeatable freshness check. The live authoritative Excerpt Database store has not been identified; Weaver's `button-weaver-internal/weaverledger/excerptRecords` collection contains 664 Weaver records and is not the full 47,461-row local source library.
