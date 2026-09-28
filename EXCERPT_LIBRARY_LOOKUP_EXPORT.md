# Excerpt Library Lookup and Export Contract

Updated: 2026-09-28

## Source and scope

- Local operational snapshot: `data/excerpt_library.db`, table `excerpt_entries` (47,461 source rows at verification). `excerpt_sources` supplies import provenance.
- Local normalized snapshot: `data/excerpt_library_normalized.db`, table `excerpts` (32,666 canonical rows at verification). Multiple source rows can map to one canonical excerpt, and non-excerpt content can be peeled off.
- These are local snapshots. Their presence and file dates do not establish that a live cloud database or downstream Google Sheet is current.
- Never claim a record is absent from the excerpt database based only on a partial Sheet range or the canonical export. Search the complete operational source row set for coverage questions.

## Full row lookup

Run from this repository:

```bash
python3 lookup_excerpt_library.py 'Tragedy and silence have the exact same address' --author 'Rudy Francisco'
```

On another machine, download `operational_excerpt_index.csv`, `lookup_excerpt_library.py`, and its normalization dependency `excerpt_library.py` from the current Drive folder, then run:

```bash
python3 lookup_excerpt_library.py 'Tragedy and silence have the exact same address' --index-path operational_excerpt_index.csv
```

The JSON result includes `scannedRowCount`, `totalMatches`, `returnedMatches`, `truncated`, match type, source name/kind/row, source record ID, author/book/poem, text, and excerpt hash. `exact_text` means normalized equality. `quote_within_excerpt` means the normalized query is contained in the longer excerpt. `possible_variant` is a scored, unconfirmed wording match that requires editorial review. The optional author filter narrows the search and must be omitted for a database-wide absence check. An empty result is limited to the specified snapshot and matching rules.

Verification case: the phrase above returned one source row for Rudy Francisco, source row 36138, poem title `“Complainers”`. Other Rudy rows associate close variants with *Helium* and `Complainers (NPS 2014)`. Do not silently merge those metadata variants into a single asserted row.

## Consumer export

Regenerate from the two local snapshots:

```bash
python3 export_weaver_excerpt_library.py
```

Output directory: `data/exports/weaver_excerpt_library/`.

- `weaver_excerpt_library.json` and `.csv`: one row per canonical normalized excerpt for Weaver matching, 32,666 rows in this snapshot.
- `operational_excerpt_index.csv`: every operational `excerpt_entries` row with source provenance and normalized lookup text, 47,461 rows in this snapshot. This is the complete local row index for coverage and QI lookup.
- `manifest.json`: schema version 4, generation time, row counts, source database and artifact SHA-256 checksums, file names portable within the export folder, and a disposition for every source row outside the canonical export.

Current coverage reconciliation: 47,461 operational rows with 33,080 distinct excerpt hashes; 32,666 canonical hashes; 634 operational rows / 414 distinct hashes outside the canonical export. Of those rows, 394 are peeled `COV`, 103 peeled `INT_LAYOUT`, and 137 classified `likely_non_excerpt`. Unaccounted rows: zero. The exporter fails if a source row lacks a disposition or either source database changes during generation.

The export is a generated snapshot, not a live feed. Consumers must use the `manifest.json` counts and checksums to identify a complete set. A refreshed export does not prove that Weaver, Poetry Please, a Google Sheet bridge, or another machine has loaded it. Those are separate delivery and ingestion checks.

Current Drive copy: [Excerpt Library Lookup Export v2](https://drive.google.com/drive/folders/11i6wGukLM6MtzsCRu3k1w55fGSDKC2Vg). The [earlier v1 folder](https://drive.google.com/drive/folders/1SZ_vwmNbOn0zf6VCB8k5x5qr7whjnME9) is retained for audit and lacks the new reconciliation fields. The local output directory is ignored by git; rerun the export command after source database changes and upload a new version as a new dated snapshot.

Live cloud comparison remains pending. The Weaver runtime configuration names `button-weaver-internal` Firestore as a candidate, but that does not establish it as the Excerpt Database's authoritative store. The configured `gcloud` account could not refresh its authentication token noninteractively. The Google Sheet workflow bridge is a downstream copy and was not used as proof of current cloud database contents.

## Next owner checks

1. Excerpt Database task: identify the existing authoritative live store, restore its approved authenticated read path, and compare it with this snapshot; keep the Rudy case as a regression check.
2. Weaver and QI consumers: use the canonical artifact for matching and the operational index for coverage/provenance; verify ingestion of a specific manifest version.
3. Legacy sheet cutover: remains a Weaver task and still requires a sheet-disabled end-to-end test.
