# Excerpt Workflow Redesign Plan

## Goal

Keep `New - Quote Creation Tool Database` as the review surface, but stop treating it as the source of truth.

The source of truth should remain `Excerpt Tool 1.20`.

## Core Shift

Current model:

- `Excerpt Tool 1.20` stores raw form responses
- `New - Quote Creation Tool Database` is a formula-built working view
- review decisions happen on the working view
- destructive actions on the working view do not persist cleanly

Proposed model:

- `Excerpt Tool 1.20` stores raw responses plus source-side workflow decisions
- `New - Quote Creation Tool Database` becomes a generated review queue
- review tools still run on the second tab
- exclude decisions get written back to the source tab
- excluded rows disappear from the second tab because the source-side filter stops surfacing them

## Source Tab Changes

Add these columns to `Excerpt Tool 1.20`:

- `record_id`
- `source_author`
- `source_title`
- `source_excerpt`
- `source_book_title`
- `exclude_from_quote_db`
- `exclude_reason`
- `duplicate_group_id`
- `duplicate_keep_record_id`
- `exact_pull_count`
- `approved_for_quote`
- `quote_created_qc`
- `added_to_primary_db`

## Meaning Of New Source Fields

- `record_id`: stable unique ID for each response row
- `source_author`: normalized resolved author name from either intake path
- `source_title`: normalized resolved poem title from either intake path
- `source_excerpt`: normalized resolved excerpt from either intake path
- `source_book_title`: normalized resolved book title from either intake path
- `exclude_from_quote_db`: `Y` if this row should not appear in the second tab
- `exclude_reason`: exact duplicate, bad quote, typo source issue, etc.
- `duplicate_group_id`: group label for exact duplicate clusters
- `duplicate_keep_record_id`: which source record wins that exact duplicate group
- `exact_pull_count`: how many times the same exact excerpt was submitted

## Second Tab Changes

Keep `New - Quote Creation Tool Database`, but repopulate it from source-side normalized columns.

Add these generated columns near the front:

- `Source Row`
- `Record ID`

Then populate the review tab from source fields:

- timestamp
- email
- resolved author
- resolved title
- resolved video/book field
- resolved book title
- resolved excerpt
- notes

Then keep the review columns on the second tab:

- quote approval
- typo review
- quote created / QC
- primary database added
- duplicate review outputs
- keep/delete recommendation

## Filtering Rule

Only rows where `exclude_from_quote_db != "Y"` should appear on the second tab.

That means the second tab becomes a filtered review queue, not a full mirror.

## Duplicate Workflow

1. Analyze duplicates on the second tab.
2. Select keep/exclude winners there.
3. Write those decisions back to the source tab using `Source Row` or `Record ID`.
4. Refresh the second-tab formulas.
5. Excluded rows disappear automatically.

## Why This Is Better

- decisions persist in the source of truth
- the second tab remains safe for reviewers
- deletes are replaced by reversible exclusions
- exact-pull count can stay attached to the kept source record
- row order changes in the second tab become less dangerous

## Implementation Order

1. Add source-side helper columns to `Excerpt Tool 1.20`
2. Generate stable `record_id` values
3. Replace the row-2 formulas in `New - Quote Creation Tool Database`
4. Update duplicate tools so they write exclusions upstream
5. Stop using row deletion on the second tab

## Recommendation

Do not try to preserve the current delete workflow.

Keep the review experience, but move all permanent state changes to the source tab.
