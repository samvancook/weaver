# Excerpt Review Tool Scope

## Product Goal

Replace the spreadsheet review workflow with a lightweight internal app that supports:

1. intake from form submissions
2. duplicate review and exclusion
3. approval / winner selection
4. output/export for quote graphics and downstream databases

## Why We Are Building This

The spreadsheet now works well enough to keep operations moving, but it is not a good long-term interface for:

- duplicate cluster review
- winner selection
- stateful moderation
- preserving attention signals like exact pull count
- preparing excerpts for graphics/export

## MVP Outcome

An internal review app where a user can:

- see incoming excerpt submissions
- review exact duplicates grouped together
- keep one and exclude the rest
- review overlap candidates
- approve excerpts for quote creation
- mark excerpts as used/exported

## Proposed Stack

Use the same shape as the current Poetry Please apps:

- Firebase Hosting for the frontend
- Firebase Functions for server/API logic
- Firestore for review-state storage
- optional Google Sheets sync during transition

## MVP Screens

### 1. Intake Queue

- list of excerpt submissions
- filters by author, book, poem, source type, state
- quick counts

### 2. Duplicate Review

- exact duplicate groups
- keep/exclude actions
- line-break-aware keep recommendation
- exact pull count

### 3. Overlap Review

- near-duplicate pairs/groups
- score, shared text preview, manual resolution

### 4. Approved Excerpts

- approved queue
- sortable by word count / char count
- winner selection for graphics

### 5. Export Queue

- excerpts ready for graphics
- export format for downstream tools

## MVP Data Model

### excerpt_submissions

- id
- source_submission_id
- author
- title
- book_title
- excerpt_text
- normalized_excerpt
- source_type
- notes
- created_at

### excerpt_review_state

- submission_id
- excluded
- exclude_reason
- duplicate_group_id
- keep_submission_id
- exact_pull_count
- approved_for_quote
- quote_created_qc
- added_to_primary_db
- export_status
- reviewer_notes

### overlap_matches

- left_submission_id
- right_submission_id
- overlap_score
- match_type
- status

## Migration Strategy

### Phase 1

- keep Google Form intake
- sync source sheet rows into the app
- use the app as the review interface

### Phase 2

- sync approved/excluded decisions back to the sheet if needed
- shift exports downstream from the app instead of the sheet

### Phase 3

- optionally replace Google Form intake with in-app submission

## Build Priorities

1. ingest sheet data into a stable app model
2. exact duplicate review
3. approval state + export queue
4. overlap review
5. graphics-facing export

## Near-Term Weaver Review UX

- split excerpt-library review into `Exact library matches` and `Possible library matches`
- default exact library matches to `Reject` so duplicate cleanup is faster and safer for interns
- keep bulk actions out of the current flow for now, but revisit them once the exact-match lane feels reliable

## Non-Goals For MVP

- full visual graphic creation inside the app
- replacing every spreadsheet workflow on day one
- complex role/permission systems

## Recommendation

Build this as a standalone internal tool first, but keep the stack compatible with Poetry Please so it can be folded in later if that proves useful.
