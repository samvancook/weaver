# Workflow DB Integration Plan

## Goals

1. Get the excerpt database to a point where it can support the existing workflow that currently depends on:
   - the primary excerpt spreadsheet
   - the Looker Studio report
2. Capture the current workflow and identify upgrades so we can reduce spreadsheet dependence over time.

## Operating Decision

For this workflow, the database is the system of record and the spreadsheet/report layer is now a compatibility layer.

That means:

- `excerpt_entries` is the operational excerpt source
- DB-built exports are the supported downstream handoff
- the old spreadsheet should be treated as a legacy input and comparison reference only
- Looker Studio should eventually point to DB-built exports, not an independently maintained sheet

## Current Workflow Inputs

### Current sheet

Source:
- `https://docs.google.com/spreadsheets/d/16WDniM-dosOy0rDvCMWaLGVGWOs8-CGeY_K3tuB1ui4/edit?gid=0#gid=0`

Observed tab:
- `Excerpt Database`

Observed header fields include:
- `Author`
- `New Title`
- `Title`
- `Title w/o quotes`
- `Quote`
- `Book`
- `DO NOT POST`
- `Other Info`
- `Graphic Made?`
- `Playlist Link:`
- `Bit.ly Link:`
- `Video Title`
- `File Name`
- `Link`
- `Button Author`
- `Author Quote Docs`
- `New Quote Tool`
- `When Pulled`
- `Who Pulled`
- `Quote Image Created`
- `Approved for Quote Image`
- `# of Characters`
- `USED`
- `Low Resolution File Name`
- `(Low Resolution) Direct Link to file (for easy downloading)`
- `Book Shortener`
- `File Type Helper`
- `File Type`
- `Use For Social`
- `Use for Ads`
- `SPECIFIC QUOTE`
- `Starting Line?`
- `Alternate Filename?`
- `Notes`
- `Book Release Year`
- `Poet FB Handle`
- `Poet IG Handle`
- `Full page or Excerpt`
- `Page Number`
- `Link to folder on drive`
- `Book Release Catalog`
- `Bestseller`

### Current reporting tool

Source:
- `https://datastudio.google.com/u/0/reporting/77c69751-4649-4d59-994e-11686942097f/page/p_xv1hv2nlud`

Note:
- I have the link, but not a detailed field inventory from the report itself yet.
- Before we swap data sources, we should explicitly inventory the dimensions, metrics, and filters the report uses.

## Current Database Status

### Main DB

File:
- `data/excerpt_library.db`

Current state as of 2026-06-23:
- `47,006` curated excerpt rows in `excerpt_entries`
- `8,721` raw spreadsheet rows in `raw_import_rows`
- live Weaver sync cursor present in `data/weaver_approved_sync_state.json`
- last live sync cursor: `2026-06-16T21:00:57.000Z`

### Normalized DB

File:
- `data/excerpt_library_normalized.db`

Current state after latest rebuild:
- canonical excerpt layer for exports and matching

### Current live Weaver sync

Script:
- `sync_weaver_approved_feed.py`

Current behavior:
- pulls approved records from live Weaver export
- ingests idempotently
- stores state cursor
- rebuilds normalized DB
- refreshes Weaver export artifacts

This means the database is no longer static; it can be kept current from Weaver.

## What “Up To Speed” Means

For the DB to support the existing workflow, it needs to satisfy two conditions:

1. Data compatibility
- The DB must expose the fields the current sheet/report expects.

2. Operational trust
- The DB must be updated reliably enough that the workflow does not need the old sheet as the source of truth.

## Recommended Path

## Phase 1: Compatibility Layer

Build a stable export/view that mirrors the existing spreadsheet workflow fields closely enough that the current Looker Studio report and spreadsheet-based habits can keep working.

Deliverable:
- one “workflow export” artifact from the DB

Recommended shape:
- one row per excerpt workflow record
- preserve the spreadsheet-friendly columns people actually use
- derive missing fields where possible from the DB

Likely source tables:
- `excerpt_entries`
- `raw_import_rows`
- `raw_excerpt_links`
- normalized/export layers where needed

## Phase 2: Field Mapping Inventory

Map:
- current spreadsheet columns
- current Looker Studio fields
- source DB fields
- derived/calculated fields

This should identify:
- direct mappings
- renamed mappings
- fields still missing from the DB
- fields that should be deprecated

## Phase 3: Workflow Export

Create a script that emits a spreadsheet-compatible CSV/JSON for the current workflow.

Candidate output:
- `data/exports/workflow_excerpt_database/workflow_excerpt_database.csv`

The goal is not to redesign the workflow yet.
The goal is to let the current workflow read from the DB with minimal disruption.

## Immediate Cleanup Path

1. Keep the old sheet for comparison/history only.
2. Regenerate the workflow export from the DB whenever sync runs.
3. Validate a few key authors/books/lane filters against the export.
4. Move the report and any review tooling to the DB-backed export.
5. Stop making manual workflow corrections in the legacy sheet once parity is acceptable.

## Phase 4: Report Validation

Before replacing the current source:
- compare record counts
- compare a few sample authors/books/titles
- compare key status fields
- compare a few report filters and totals

Only after that should we point the report to the DB-backed export.

## Phase 5: Workflow Upgrade

Once the DB is proven in the current workflow, improve the workflow itself.

Likely upgrade targets:
- stop using the primary excerpt spreadsheet as the operational source of truth
- split raw/archive concerns from canonical excerpt concerns
- give status views for:
  - approved but unlinked
  - rejected
  - excerpt text present but not canonicalized
  - graphics eligibility / quote-image status
- reduce manual spreadsheet edits by producing exports instead

## Immediate Development Needs

To make DB updates from Weaver simple and automatic, we still need:

1. Scheduled sync execution
- run `sync_weaver_approved_feed.py` automatically

2. Sync monitoring
- log failures and last successful sync

3. Workflow export build
- a dedicated export shaped for the current sheet/report workflow

4. Looker field inventory
- explicit list of what the current report expects

## Suggested Next Implementation

Build:
- `export_workflow_excerpt_database.py`

Purpose:
- create a DB-backed export shaped like the current `Excerpt Database` sheet

That is the safest next step because it gets the database usable in the current workflow before we try to redesign the workflow itself.

## Big-Picture End State

The stable end state should be:

- Weaver approved excerpts flow into the DB automatically
- the DB rebuilds canonical/normalized layers automatically
- exports refresh automatically for downstream tools
- Looker/reporting reads DB-built exports
- the old spreadsheet is retired from active use

## 2026-07-15 Charm City And Source-Of-Truth Checkpoint

### Boundary now implemented

The operating flow is:

`Weaver intake and approval -> raw_import_rows + excerpt_entries -> normalized DB -> exports -> reporting`

Firestore runs alongside this flow and owns graphics lifecycle state only. It is not an excerpt-content database.

Responsibilities:

- Excerpt Gathering sheet: intake and temporary review surface
- Weaver: excerpt collection, review, and approval
- `raw_import_rows`: immutable-enough archive of source rows and approved Weaver payloads
- `excerpt_entries`: operational excerpt truth
- `excerpt_library_normalized.db`: canonical and deduplicated truth
- `data/exports/`: generated delivery artifacts only
- Firestore `weaverledger`: graphics request, production, QC, rework, and handoff lifecycle
- Codex Shared Drive: documentation, reconciliation reports, and published artifacts; not an editable excerpt database

### Implemented in this checkpoint

- Weaver approved-export records now carry `sourceEvent` separately from `bookTitle`.
- Video excerpt exports read video author, poem, excerpt, and real book fields correctly.
- Video excerpts may have no book without substituting an event name.
- `sync_weaver_approved_feed.py` remains the supported ingest path.
- Approved Weaver payloads are archived idempotently in `raw_import_rows` before operational upsert.
- Approved records are upserted idempotently into `excerpt_entries`.
- Successful sync rebuilds `excerpt_library_normalized.db` and refreshes exports.
- Weaver revision `weaver-00342-g87` deployed the corrected export contract.
- Code checkpoint: `e021336` (`Separate excerpt event metadata from book identity`).

### BPL Charm City 2026 reconciliation

Live source set:

- 132 video excerpt rows
- 52 submitted by `ljohnson@buttonpoetry.com`
- 79 submitted by `sdrayton@buttonpoetry.com`
- 1 submitted by `sam@buttonpoetry.com`
- 0 rows remain labeled `Baltimore`

Database reconciliation after the 2026-07-15 sync:

- 123 approved source rows archived from Weaver
- 118 operational excerpts after matching/deduplication
- 5 source rows matched or collapsed into existing operational content
- 9 source rows are not yet approved and remain outside operational truth
- 0 operational rows use `BPL Charm City 2026` as `book_title`
- normalized database rebuilt successfully
- Weaver JSON, CSV, and manifest exports refreshed successfully

The five matched rows remain individually preserved in `raw_import_rows`. Operational and canonical layers should not create duplicate excerpt content merely to preserve source-row identity.

### Remaining roadmap

1. Archive the nine unapproved Charm City rows directly from the intake source without adding them to `excerpt_entries`.
2. Review the nine rows in Weaver and explicitly approve, reject, or request correction.
3. Add automated scheduled execution for `sync_weaver_approved_feed.py`.
4. Add monitoring for last successful sync, fetched count, cursor movement, and failures.
5. Add regression coverage enforcing:
   - `sourceEvent` never replaces `bookTitle`
   - approved Weaver records enter both archive and operational layers
   - unapproved records do not enter operational truth
   - Firestore remains graphics lifecycle storage only
6. Build and validate the dedicated workflow/reporting export.
7. Inventory Looker Studio fields and filters.
8. Move Looker and workflow consumers to DB-generated exports.
9. Publish a concise source-of-truth document and reconciliation report in the Codex Shared Drive.
10. Remove or supersede contradictory older instructions.
11. Confirm archive coverage and downstream parity, then make the Gathering sheet archive-only/read-only.

### Pause safety

Work can pause after this checkpoint. Until the remaining tasks are completed:

- do not delete, retire, or make the Gathering sheet inaccessible
- do not treat the nine unapproved rows as operational excerpts
- do not point reporting away from its current source until DB-export parity is verified
- do not create another excerpt store or direct Firestore excerpt ingestion path
