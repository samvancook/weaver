# Excerpt Management

This workspace now includes a first-pass CSV utility for surfacing duplicate and high-overlap excerpts.

## Excerpt Library Database

There is now also a SQLite-backed excerpt library for turning a large spreadsheet export into a cleaner local database that Weaver can query.

Import a CSV export:

```bash
python3 import_excerpt_library.py /path/to/excerpts.csv --replace-source
```

This creates `data/excerpt_library.db` with:

- `excerpt_sources` for import provenance
- `excerpt_entries` for normalized excerpt rows plus searchable metadata

Once that database exists, `catalog_validate.py` will also attach `libraryExcerptMatch` to validation results so Weaver can flag excerpts that already exist in your imported library.

### 1. Exact-Match Batch Report

Use this when you want a strict report of which candidate rows already exist exactly in the library:

```bash
python3 exact_excerpt_report.py /path/to/candidates.csv --output-csv /tmp/exact_matches.csv
```

### 2. Reusable Matcher

Use this when you want a reusable JSON or CSV matcher that flags exact matches and near matches:

```bash
python3 match_excerpt_candidates.py /path/to/candidates.json
```

Or with CSV input:

```bash
python3 match_excerpt_candidates.py /path/to/candidates.csv --output-csv /tmp/match_results.csv
```

### 2a. Accepted Excerpt Bridge

Use this when you want to turn accepted rows from the old tool or Weaver into a clean report of:

- already in the excerpt library
- new candidates to add to the excerpt library

It supports both:

- explicit Weaver accepts: `excerpt_review_decision = ACCEPT`
- legacy accepts: `approved_for_quote = Y` when no explicit decision exists

Example:

```bash
python3 accepted_excerpt_bridge.py /path/to/excerpt-tool-export.csv \
  --output-csv /tmp/accepted_excerpt_bridge.csv \
  --new-candidates-csv /tmp/new_excerpt_candidates.csv
```

This gives you a catch-up bridge between:

- accepted excerpts in the working sheet
- the local excerpt library database
- the future Poetry Please import layer

### 3. Cleanup Report

Use this to profile duplicate clusters, likely non-excerpt rows, and blank-book concentration:

```bash
python3 excerpt_cleanup_report.py --top 25
```

### 4. Build A Deduped Database

Use this to build a cleaned database that collapses exact duplicate excerpt text into canonical rows while preserving all original occurrences:

```bash
python3 build_deduped_excerpt_library.py
```

This creates `data/excerpt_library_deduped.db` with:

- `canonical_excerpts`
- `canonical_excerpt_occurrences`
- `v_canonical_excerpt_summary`

Key fields:

- `pull_count`: how many times the excerpt was pulled before dedupe
- `primary_author`, `primary_book_title`, `primary_poem_title`: preferred merged values
- `author_values_json`, `book_values_json`, `poem_title_values_json`: preserved variant values
- `has_author_conflict`, `has_book_conflict`, `has_poem_title_conflict`: flags for groups where metadata drifted

The safe workflow is to treat `canonical_excerpts` as the cleaned layer and `canonical_excerpt_occurrences` as the audit trail for every original row that was merged into it.

### 4a. Classify Non-Excerpt Rows First

Before deduping, classify raw imported rows into:

- `likely_excerpt`
- `likely_non_excerpt`
- `needs_review`

Run:

```bash
python3 classify_excerpt_rows.py --json-summary
```

This writes:

- `data/excerpt_row_classification.csv`

The intended workflow is:

1. classify rows
2. quarantine `likely_non_excerpt`
3. manually inspect `needs_review`
4. dedupe only the excerpt-safe subset

### 5. Build A Normalized Excerpt Database

Use this to make excerpts the primary records and attach `QI` / `INT` / `COV` rows as linked assets instead of treating them as competing content types:

```bash
python3 build_normalized_excerpt_database.py
```

This creates `data/excerpt_library_normalized.db` with:

- `excerpts`
- `source_rows`
- `excerpt_assets`
- `peeled_off_rows`

Key ideas:

- each row in `excerpts` is one underlying excerpt
- each raw imported row is preserved in `source_rows`
- `QI` rows become linked rows in `excerpt_assets`
- peeled-off categories like `COV` are moved into `peeled_off_rows`
- `has_qi_asset` and `qi_asset_count` live on the parent excerpt row
- additional excerpt-level QI rollups include:
  - `qi_linked_asset_count`
  - `qi_approved_count`
  - `qi_graphic_made_count`

This is the first step toward the model:

- underlying excerpt
- flag that it has a quote image
- file/link metadata for that quote image

Current peeled-off categories inside the normalized DB:

- `COV`
- `INT_LAYOUT`

### 5a. Export Peeled-Off Content

Use this to export content categories that are being migrated out of the excerpt tool:

```bash
python3 export_peeled_off_content.py
```

Current peeled-off export:

- from imported SQLite library:
  - `COV` -> `data/peeled_off/from_db/cov_rows.csv`
  - `ART` -> `data/peeled_off/from_db/art_rows.csv`
- from raw CSV source:
  - `COV` -> `data/peeled_off/from_raw_csv/cov_rows.csv`
  - `ART` -> `data/peeled_off/from_raw_csv/art_rows.csv`

See [PEELED_OFF_CONTENT.md](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/PEELED_OFF_CONTENT.md) for the running list.

### 6. Ingest Approved Weaver Excerpts

Use this to ingest approved excerpts coming from Weaver into the core excerpt library:

```bash
python3 ingest_weaver_approved_excerpts.py /path/to/weaver_approved.json
```

The JSON payload can be either a list of records or an object with a `records` array.

Expected fields per record:

- `recordId`
- `sourceRow`
- `author`
- `bookTitle`
- `title`
- `excerptText`

Example:

```json
{
  "records": [
    {
      "recordId": "weaver-123",
      "sourceRow": 101,
      "author": "Test Author",
      "bookTitle": "Test Book",
      "title": "Test Poem",
      "excerptText": "This is a brand new approved excerpt from Weaver."
    }
  ]
}
```

Ingest behavior:

- writes to `data/excerpt_library.db`
- stores records under source `weaver://approved`
- upserts by `recordId` when present
- preserves the full original record JSON in `metadata_json`

After ingest, rebuild the normalized DB:

```bash
python3 build_normalized_excerpt_database.py
```

### 5b. Reconcile Approved Quote Images To Existing Graphics

Use this to build a catch-up report for approved `QI` rows and identify where a graphic already exists:

```bash
python3 reconcile_qi_graphics.py \
  --output-csv data/qi_graphic_reconciliation.csv \
  --unmatched-csv data/qi_graphic_reconciliation_unmatched.csv \
  --sheet-update-csv data/qi_graphic_reconciliation_sheet_updates.csv
```

The output buckets approved rows into:

- `direct_on_approved_row`: the approved row itself already has a filename/link
- `sibling_qi_asset_same_excerpt`: the approved row has no direct file, but the same normalized excerpt has another linked `QI` row
- `no_detected_graphic`: no linked `QI` asset was found for that approved excerpt

Helpful output columns:

- `recommendMarkGraphicMade`: likely safe rows to mark as already made
- `recommendBackfillLinkToApprovedRow`: rows where the graphic appears to exist on a sibling `QI` row and should be linked back
- `matchedAssetsJson`: audit trail of the candidate linked assets used for the match
- `sheet-update-csv`: compact row-keyed export for pushing recommended file/link/graphic-made values back upstream

## Excerpt Review App

Weaver is now the active internal review app for this project.

Production Weaver runs from:

- [public/index.html](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/public/index.html)
- [public/app.js](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/public/app.js)
- [public/style.css](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/public/style.css)
- [server.mjs](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/server.mjs)

Current direction:

- form input
- app-based review
- app-based graphics creation support
- app-based QC
- browser-direct Google Sheet export where needed

### Transitional P.I.G. Writeback

Weaver now has a transitional completion intake path for P.I.G.-generated graphics:

- `POST /api/pig/completed-graphics`
- completed graphics are stored in the dedicated `PIG - Completed Graphics` sheet lane
- those returned graphics are surfaced inside `Graphics QC`

Expected completion payload shape:

```json
{
  "completions": [
    {
      "completionId": "pig-123",
      "requestId": "weaver:20250415-123",
      "author": "Author Name",
      "poemTitle": "Poem Title",
      "bookTitle": "Book Title",
      "quoteText": "Excerpt text",
      "assetUrl": "https://drive.google.com/...",
      "assetPreviewUrl": "https://lh3.googleusercontent.com/...",
      "sourceRecordId": "20250415120000-123",
      "sourceSheetRow": 456,
      "productionNotes": "Optional note from P.I.G.",
      "completedAt": "2026-04-15T18:00:00Z",
      "sourceTool": "P.I.G."
    }
  ]
}
```

This is intentionally transitional. The roadmap still points toward a Weaver-owned database replacing this sheet lane later.

### Transitional P.I.G. Read Source

P.I.G. can now read live open graphics requests directly from Weaver:

- `GET /api/pig/graphics-request-books?filter=all`
- `GET /api/pig/graphics-request-books?filter=current_titles`
- `GET /api/pig/graphics-requests?filter=all`
- `GET /api/pig/graphics-requests?filter=current_titles`
- `GET /api/pig/graphics-requests?filter=current_titles&bookTitle=Coin%20Laundry%20at%20Midnight`
- `GET /api/pig/graphics-request-records?filter=current_titles&bookTitle=Coin%20Laundry%20at%20Midnight` for raw row-level queue records

The main `graphics-requests` endpoint now returns grouped request objects. Each request includes:

- `graphicsRequestId`
- `bookTitle`
- `poemTitle`
- `author`
- `requestStatus`
- `completionCount`
- `latestCompletedAt`
- `assetUrl`
- `assetPreviewUrl`
- `queueSheetRows`
- `excerpts[]`

The raw `graphics-request-records` endpoint still returns row-level queue records when needed.

This is the live Weaver-derived source P.I.G. should prefer over stale local snapshots.

### Poetry Please Handoff Feed

Poetry Please can now read QC-approved graphics from a dedicated Weaver handoff feed:

- `GET /api/poetry-please/qc-approved-graphics/books`
- `GET /api/poetry-please/qc-approved-graphics`
- `GET /api/poetry-please/qc-approved-graphics?bookTitle=Stunt%20Water%3A%20The%20Work%20of%20Buddy%20Wakefield`

This feed returns only graphics that:

- came back from P.I.G.
- were QC-approved in Weaver
- have a usable asset link

Each record currently includes:

- `sourceCompletionId`
- `sourceRequestId`
- `title`
- `author`
- `book`
- `excerpt`
- `driveLink`
- `previewUrl`
- `imageType`
- `completedAt`
- `qcApprovedAt`
- `qcNote`
- `productionNotes`
- `sourceTool`

`bookLink` and `releaseCatalog` are intentionally blank for now so Poetry Please can fill them from its own catalog mapping layer.

## Weaver Runtime Database

The first internal Weaver-owned database slice is now defined for the graphics pipeline.

Files:

- [db/weaver_runtime_schema.sql](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/db/weaver_runtime_schema.sql)
- [weaver_runtime_db.py](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/weaver_runtime_db.py)
- [init_weaver_runtime_db.py](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/init_weaver_runtime_db.py)
- [WEAVER_DATABASE_PLAN.md](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/WEAVER_DATABASE_PLAN.md)

Initialize a local runtime DB:

```bash
python3 init_weaver_runtime_db.py
```

This locks the schema and gives us a migration target for:

- graphics requests
- P.I.G. completions
- QC reviews
- Poetry Please handoffs

It is not the final production persistence layer for Cloud Run. The intended production target is Cloud SQL Postgres, with the spreadsheet retired in phases.

## Apps Script Sync

The Google Apps Script project is now connected with `clasp`.

- Root `Code.js`: Excerpt Update Tool / sheet-bound maintenance script
- Root `appsscript.json`: Excerpt Update Tool manifest
- Older local reference file: `google_apps_script.gs`

Useful commands:

```bash
npm run clasp:pull
npm run clasp:push
```

Important distinction:

- `npm run clasp:pull` / `npm run clasp:push` target the root Apps Script project from `.clasp.json`
- that root synced project is **not** the Cloud Run-hosted Weaver app
- changes in `public/index.html`, `public/app.js`, `public/style.css`, or `server.mjs` require a Cloud Run deploy to affect `weaver.buttonpoetry.com`
- changes in `Code.js` affect the Excerpt Update Tool Apps Script, not hosted Weaver

`google_apps_script.gs` is intentionally ignored by `clasp` so we do not accidentally push the wrong file.

## What it adds

Run the script against a CSV file and it will append these columns:

- `excerpt_word_count`
- `excerpt_character_count`
- `duplicate_status`
- `matched_row_number`
- `overlap_score`
- `shared_token_count`
- `matched_excerpt_preview`

`duplicate_status` is currently:

- `exact_duplicate` when two excerpts normalize to the same text
- `high_overlap` when two excerpts are very similar but not identical

## Expected input

The script tries to find one of these column names automatically:

- `Enter the Quote`
- `Quote`
- `Excerpt`
- `Text`

If your sheet uses a different name, pass `--quote-column`.

## Run it

```bash
python3 excerpt_analyzer.py input.csv output.csv
```

If you want stricter or looser overlap detection:

```bash
python3 excerpt_analyzer.py input.csv output.csv --near-duplicate-threshold 0.80
```

Lower thresholds catch more possible overlap but will create more false positives.
