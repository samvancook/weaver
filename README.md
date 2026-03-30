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

### 3. Cleanup Report

Use this to profile duplicate clusters, likely non-excerpt rows, and blank-book concentration:

```bash
python3 excerpt_cleanup_report.py --top 25
```

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
