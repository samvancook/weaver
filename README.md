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

This is the first step toward the model:

- underlying excerpt
- flag that it has a quote image
- file/link metadata for that quote image

### 5a. Export Peeled-Off Content

Use this to export content categories that are being migrated out of the excerpt tool:

```bash
python3 export_peeled_off_content.py
```

Current peeled-off export:

- `COV` -> `data/peeled_off/cov_rows.csv`
- `ART` -> `data/peeled_off/art_rows.csv`

See [PEELED_OFF_CONTENT.md](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/PEELED_OFF_CONTENT.md) for the running list.

## Excerpt Review App

This repo now also includes the first scaffold for a Firebase-style `Excerpt Review Tool`:

- [PRODUCT_SCOPE.md](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/PRODUCT_SCOPE.md)
- [firebase.json](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/firebase.json)
- [functions/index.js](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/functions/index.js)
- [public/index.html](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/public/index.html)

The current direction is to move from:

- form input
- spreadsheet review
- manual export

to:

- form input
- app-based review
- app-based export

## Apps Script Sync

The Google Apps Script project is now connected with `clasp`.

- Canonical synced script: `Code.js`
- Apps Script manifest: `appsscript.json`
- Older local reference file: `google_apps_script.gs`

Useful commands:

```bash
npm run clasp:pull
npm run clasp:push
```

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
