# Peeled-Off Content

This file tracks content categories that have been intentionally removed from the excerpt-focused database/tooling and prepared for migration elsewhere.

## Current List

### `COV`

- Status: peeled off from excerpt normalization
- Meaning: cover/interior/packaging style asset rows, not underlying excerpts
- Why removed: these rows distort excerpt pull counts, duplicate groups, and excerpt-level matching
- Export from imported SQLite library: `data/peeled_off/from_db/cov_rows.csv`
- Export from raw CSV: `data/peeled_off/from_raw_csv/cov_rows.csv`
- Manifests:
  - `data/peeled_off/from_db/manifest.json`
  - `data/peeled_off/from_raw_csv/manifest.json`
- Current row count in imported SQLite library: `394`
- Current row count in raw CSV: `775`

### `ART`

- Status: peeled off from excerpt normalization
- Meaning: art / UGC / fan-art style asset rows, not underlying excerpts
- Why removed: these rows are book-linked assets with no excerpt text and do not belong in excerpt matching or excerpt counts
- Export from imported SQLite library: `data/peeled_off/from_db/art_rows.csv`
- Export from raw CSV: `data/peeled_off/from_raw_csv/art_rows.csv`
- Manifests:
  - `data/peeled_off/from_db/manifest.json`
  - `data/peeled_off/from_raw_csv/manifest.json`
- Current row count in raw CSV: `3`
- Current row count in imported SQLite excerpt library: `0`
- Note: these rows never entered the excerpt SQLite import because they have blank quote text

## Future Candidates

- `ART`
- some `INT` rows that are clearly layout/asset instructions rather than excerpt-derived content
