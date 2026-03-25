# Excerpt Management

This workspace now includes a first-pass CSV utility for surfacing duplicate and high-overlap excerpts.

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
