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

Current direction:

- Cloud Run for the live Weaver frontend and API
- Google Sheets as the operational source during transition
- browser-direct Google Sheets export for user-owned exports
- a future Weaver-owned database for app-native workflow state like QC
- recommended production database target: Cloud SQL Postgres, with local SQLite modeling during migration

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
- keep exact library matches recommendation-only, not auto-selected
- keep bulk actions out of the current flow for now, but revisit them once the exact-match lane feels reliable

## Current Workflow Direction

- Weaver remains the review and correction layer
- the excerpt library database becomes the canonical approved excerpt layer
- Weaver should become the canonical source for graphics-creation requests
- P.I.G. should consume graphics-creation requests from Weaver-owned data, not from an export-only spreadsheet
- Poetry Please should eventually read from that canonical excerpt layer instead of from ad hoc sheet state

## Intended End-To-End Workflow

The clearer long-term Weaver flow is:

1. Excerpt gathering
2. Excerpt approval for graphics
3. Graphic creation support
4. Graphic QC
5. Downstream publishing handoff

Notes:

- `Excerpt gathering` now needs to become a real Weaver module, with a path to stand alone later
- `Excerpt approval` is today’s accept / reject / needs correction review flow
- `Graphic creation support` should only show graphics that still need to be made and should ultimately populate a Weaver-owned graphics-request database
- `P.I.G.` should read from that graphics-request database and create the graphics
- `Graphic QC` should only show graphics that already exist, ideally with direct Drive links
- `Graphic QC` should receive completed graphics back from P.I.G. for approval / correction / recreate decisions
- once a graphic passes QC, the desired downstream state is to hand it off to Poetry Please automatically
- non-approval QC outcomes should split cleanly:
  - `Correct and recreate` should create a new active rework request for P.I.G. and reappear in Weaver's creation queue for parity
  - `Mismatched graphic` should route into a separate pairing workflow, not back into normal creation
  - `Final reject` should leave the active queues and remain available only in audit/history
- the first concrete schema and migration plan now live in [WEAVER_DATABASE_PLAN.md](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/WEAVER_DATABASE_PLAN.md)

### Excerpt Gathering Module

This should start life as a Weaver module and later be able to stand on its own as a dedicated app.

Primary job:

- replace the current excerpt gathering Google Form
- capture excerpts cleanly into the Weaver ecosystem
- preserve enough book / source metadata to feed review, correction, and eventual excerpt-database sync

Victory condition:

- users can gather excerpts from ebooks inside the tool
- enter excerpts manually, in bulk, or by selecting / highlighting text from the manuscript
- the tool attaches the correct book metadata automatically
- each excerpt becomes a durable record that feeds into Weaver review and then into the larger excerpt database

Phased build:

1. Phase 1: replicate the current excerpt gathering form inside Weaver
   - support the current form's three jobs:
     - add a quote from a book
     - add a quote from a video
     - fix a quote in one of the quote tools
   - preserve the current required metadata and formatting expectations from the SOP
   - make the output land in the same operational ecosystem the current form feeds
   - split those jobs into clearer visible lanes instead of stacking them together in one screen

2. Phase 2: ebook-native gathering
   - open or attach the relevant ebook / PDF inside the tool
   - allow single-entry, bulk-entry, and highlight-to-capture gathering
   - let users select text and send it directly into excerpt creation with one action
   - auto-bind the excerpt to the active book / source metadata
   - support poem-to-poem navigation in reading order for EPUB-backed books
   - optionally capture lightweight full-poem reactions during the same reading pass

Near-term product stance:

- Phase 1 is enough to count as a successful first ship
- Phase 2 is the fuller long-term win condition
- we should not block the replacement on the perfect ebook-native interaction if we can ship the current form workflow first
- the working implementation + next-step detail now live in [EXCERPT_GATHERING_MODULE_PLAN.md](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/EXCERPT_GATHERING_MODULE_PLAN.md)
- near-term follow-up: confirm whether quote-fix intake should remain in this module or live entirely in the existing correction workflow
- near-term follow-up: confirm the long-term canonical home for author, book, and poem metadata instead of leaving intake suggestions split across sources
- near-term follow-up: use the legacy Google Form suggestion lists as the fallback for non-EPUB book intake until the canonical metadata source is decided
- near-term follow-up: design EPUB navigation so a curator can move through poems in sequence instead of treating EPUB context as a one-poem popup
- near-term follow-up: use the publishing-order sheet as the source for forthcoming / non-EPUB titles until the stronger metadata pipeline is ready
- near-term follow-up: decide whether the long-term metadata pipeline should come from a regular database scrape, an intentional database push, or both

### Target Weaver -> P.I.G. -> Weaver Shape

Long-term target:

1. Weaver approval creates a graphics-request record in a Weaver-owned database.
2. P.I.G. reads the open graphics-request records directly.
3. P.I.G. creates the graphic and writes back completion metadata:
   - graphic asset location
   - created timestamp
   - source request id
   - any production notes
4. Weaver surfaces those completed graphics in `Graphics QC`.
5. QC decisions live in Weaver-owned state, not spreadsheet columns.
6. QC-approved graphics become eligible for Poetry Please handoff.
7. Poetry Please should own the final bucket upload + Firestore write for approved graphics, even when Weaver is the system that triggers the handoff.

Near-term implications:

- the spreadsheet can remain a transition surface, but it should stop being the long-term source of truth for graphics creation requests
- `Needs graphics` should eventually populate database records, not just export rows
- P.I.G. integration should be designed around stable ids instead of title/text-only matching
- as a transitional step, P.I.G. completion writeback can land in a dedicated Weaver-owned sheet lane and still surface in `Graphics QC`
- once a graphic is returned to Weaver QC, it should leave the creation queue immediately, even before QC approval
- once a graphic has any QC decision, it should leave the QC queue and stop presenting as unresolved work
- when a QC decision is a true rework request, Weaver should generate a fresh downstream request identity so P.I.G.'s suppression logic does not hide the rework item
- Weaver should surface structured rework context to P.I.G.:
  - reject reason
  - metadata issue
  - aesthetic concern
  - reviewer note

### P.I.G. Priority Order

To keep P.I.G. and Weaver in sync, the order should be:

1. Fix the data pipe
   - define the Weaver-owned source P.I.G. should read
   - make refresh/rebuild behavior predictable
   - stop relying on stale local snapshots as the main source of truth

2. Make P.I.G. consume that source cleanly
   - needs graphics
   - request ids
   - existing asset links
   - QC-ready completions

3. Add write-back
   - completed asset links
   - generated asset metadata
   - QC-return status

Current recommendation:

- prioritize the graphics-request data layer before more P.I.G. UI work
- prefer a direct Weaver-derived source over another export-only intermediate if we can
- treat stale snapshot behavior as a systems bug, not a UI bug

### QI Cleanup Catch-Up (Paused)

We started a live pass to use the QI database to remove already-made excerpts from Weaver's `Needs graphics` lane, then paused before applying writebacks.

What we confirmed:

- the right live source is Weaver's current graphics queue, not the stale local cleanup CSV
- the direct `qi_image_database.csv` exact-match pass is too skinny on the current queue
- the normalized QI database is a better source for current cleanup work
- a conservative linked-asset pass found a small set of live candidates worth revisiting first

Where to resume:

1. rebuild the current live candidate set from Weaver's graphics queue feed
2. match against the normalized QI database, preferring rows with linked assets over weaker status-only signals
3. verify each proposed source row against the live `New - Quote Creation Tool Database`
4. only then write `AQ = Y` so those rows fall out of `Queue - Needs Graphics`

Guardrail:

- keep this as a conservative cleanup pass until the spreadsheet-to-database migration is farther along

### Graphic QC Decision Set

The current intended QC decisions are:

- `Approve`
  Use when the existing graphic is correct and ready to move forward.
- `Mismatch`
  Use when the image and the excerpt do not match each other.
- `Correct`
  Use when the graphic is basically right, but the text or metadata needs fixing.
- `Recreate`
  Use when the graphic should be remade because the design or readability is not good enough.

Working rules:

- `Mismatch`, `Correct`, and `Recreate` should require or strongly encourage a note
- `Approve` should be the fastest path
- a future downstream step should allow `Approve` to hand QC-passed graphics to Poetry Please automatically

## Accepted Excerpt Bridge

### Immediate Goal

Make sure accepted excerpts can move cleanly from:

1. the old tool
2. Weaver
3. into the excerpt library database
4. and then on to Poetry Please

### Eventual Must-Have Outcome

We need a complete pathway from Weaver and the old excerpt collection tool into the excerpt library database, even if we do not finish that migration immediately.

That means:

- Weaver-approved excerpts must be able to land in the excerpt library database reliably
- accepted excerpts from the old excerpt collection tool must also be brought forward
- the migration must include a catch-up / reconciliation pass so current accepted excerpts are not stranded in either tool
- we need a verification step proving that the current accepted excerpt universe has been accounted for in the excerpt library database

Success condition:

- no currently accepted excerpt lives only in Weaver or only in the old tool without a confirmed path into the excerpt library database

### Forward Path

- accepted excerpts in Weaver should be eligible for sync into the excerpt library database
- the sync should check whether the accepted excerpt already exists in the library
- exact-existing excerpts should be marked as already captured
- only new candidates should be prepared for import downstream

### Catch-Up Path

Once the tool is tested and catch-up review is done:

- gather historical accepted excerpts from the old tool
- gather historical accepted excerpts from Weaver
- compare both sets against the excerpt library
- import any accepted excerpts that are still missing
- run a combined reconciliation across Weaver + old tool + excerpt library so we can confirm coverage of all currently accepted excerpts

### First Tooling Step

- build a bridge report from accepted sheet rows into the excerpt library
- classify each accepted row as:
  - `already_in_library`
  - `new_candidate`
- use that report to drive the first backfill into the excerpt library

### Migration Work We Still Need

1. Build a durable Weaver -> excerpt library database sync path
2. Define the old excerpt collection tool catch-up dataset
3. Reconcile both sources against the current excerpt library
4. Produce a verification report showing:
   - already present in database
   - newly imported
   - still unresolved / needs review
5. Do not consider this migration complete until that verification report exists

## Roadmap

### Now

- stabilize hosted Weaver review behavior
- keep duplicate review intern-safe without hidden defaults
- keep the excerpt library match visible and understandable
- make comparison popups easy to use during review
- keep P.I.G. reading from a live Weaver-derived source instead of stale snapshots

### Next

- generate accepted-excerpt bridge reports from sheet exports
- define the rule for when an accepted Weaver excerpt is ready to enter the excerpt library
- produce a clean list of net-new accepted excerpts for import
- keep the main review queue limited to a current-season allowlist:
  - current 2026 titles
  - `A Choir of Honest Killers`
  - upcoming fall titles added in one config place
- start the backend migration by moving read-only review endpoints behind Cloud Run first:
  - `pendingRecords`
  - `excerpts`
  - `correctionBooks`
  - `corrections`
  - `graphicsBooks`
  - `graphicsRecords`
- split the current graphics module into two clearly named lanes:
  - `Graphics Creation Queue`
  - `Graphics QC Queue`
- keep already-created graphics out of the creation queue whenever they can be matched from Drive + the QI database
- expose Drive links for QC-visible graphics inside Weaver
- add QC decision storage and UI for:
  - `Approve`
  - `Mismatch`
  - `Correct`
  - `Recreate`
- complete the final backend migration step for interactive sheet writes:
  - move QC saves off the legacy Apps Script action path
  - keep normal queue reads/saves working through Cloud Run while this is built

### Weaver Next Steps

1. Make `Needs graphics` a first-class Weaver data model
   - keep the live `/api/pig/graphics-requests` feed stable
   - make `graphicsRequestId` the required id for downstream work
   - reduce duplicate request ambiguity where multiple rows share one poem/request shape

2. Separate transitional and long-term storage clearly
   - keep `PIG - Completed Graphics` working as a transitional return lane
   - move QC state out of spreadsheet columns into Weaver-owned storage next
   - define the first real graphics-request database table after that

3. Tighten Graphics QC around returned assets
   - keep inline previews working
   - improve asset/source labeling so P.I.G.-returned work is obvious
   - preserve a clear audit trail from request -> completion -> QC decision

4. Prepare the Poetry Please handoff
   - define the exact payload a QC-approved graphic should send downstream
   - make sure the QC approval state is durable enough to drive that handoff

### Integration Next Steps

1. P.I.G. should switch to Weaver’s live request feed
   - prefer `/api/pig/graphics-request-books`
   - prefer `/api/pig/graphics-requests`
   - stop depending on stale local snapshot rebuilds as the primary source

2. P.I.G. should treat `graphicsRequestId` as the stable handshake id
   - use it in its local processing
   - return it on completion writeback

3. P.I.G. writeback should use the completion endpoint
   - `POST /api/pig/completed-graphics`
   - include asset link, preview link if available, request id, and production notes

4. Poetry Please integration should wait until QC approval is stable
   - do not bypass QC
   - handoff should happen from approved QC state, not from graphic creation completion alone

### Shared Contract We Need To Lock

- request id:
  - `graphicsRequestId`
- request source:
  - live Weaver queue feed
- completion return:
  - `POST /api/pig/completed-graphics`
- QC review target:
  - Weaver `Graphics QC`
- downstream publish target:
  - Poetry Please only after QC approval

## Stability Audit

### Current Production Truth

- `weaver.buttonpoetry.com` is the real production Weaver app
- production frontend + API shell run on Cloud Run from:
  - `public/index.html`
  - `public/app.js`
  - `public/style.css`
  - `server.mjs`
- Cloud Run now owns normal review/correction/graphics reads
- Cloud Run now owns normal review saves
- Cloud Run now owns graphics QC saves
- direct Google Sheets reads use the Cloud Run service account
- browser OAuth should only be needed for explicit Google Sheet export actions

### Still-Confusing Surfaces

1. Root `Code.js` still looks like it might be Weaver
- root `Code.js` is the Excerpt Update Tool / sheet-bound maintenance script
- confusion risk:
  - it is easy to think a `clasp push` changes hosted Weaver when it does not

2. The repo still contains stale planning language from the Firebase-first phase
- confusion risk:
  - docs can point us toward the wrong deployment model during an incident

3. Secrets and local auth artifacts can drift into the workspace
- example: local OAuth client secret downloads
- confusion risk:
  - creates avoidable security noise and makes it harder to see what is source code vs local machine state

### Immediate Cleanup Plan

1. Keep Cloud Run as the only production Weaver host.
2. Treat root `Code.js` as Excerpt Update Tool only.
3. Keep Google OAuth only for user-owned export flows.
4. Ignore local OAuth client-secret files in git.
5. Remove dead legacy routes and files as soon as their replacements are verified.

### “Before We Deploy” Checklist

1. Did we change `public/*`, `server.mjs`, or both?
2. Did we change root `Code.js`?
3. Which live surface should reflect the change?
4. What is the one marker we will check on `weaver.buttonpoetry.com` after deploy?

### Migration Status

- browser JSONP has been removed from hosted Weaver
- browser-held Apps Script URL state has been removed from hosted Weaver
- Cloud Run now owns:
  - bootstrap config
  - review queue reads
  - correction queue reads
  - graphics queue reads
  - review save requests
  - graphics QC save requests
- Cloud Run now also normalizes capitalization/spacing variants in queue book lists, including the graphics queue
- direct Google Sheets reads now work through the Cloud Run service account
- graphics export now uses user-owned browser OAuth directly against Google Sheets
- legacy Firebase scaffolding has been removed from the repo
- the in-repo Weaver Apps Script bridge has been removed from the repo
- legacy server / Apps Script export routes have been removed

### After Catch-Up

- sync accepted Weaver excerpts into the excerpt library on an ongoing basis
- expose excerpt-library presence and production status more clearly in Weaver
- validate whether accepted excerpts also exist as quote images / downstream content
- define the first Weaver-owned graphics-request table for P.I.G. handoff
- define the writeback shape P.I.G. should use when returning completed graphics for QC
- keep the transitional `PIG - Completed Graphics` lane thin and replaceable so it can be swapped for a real database later

### Later

- move the backend off Apps Script if the review workflow keeps expanding
- use the excerpt library as the cleaner canonical source for Poetry Please
- replace legacy sheet/tool dependencies with clearer database-backed flows
- move graphics QC state out of the spreadsheet and into Weaver-owned storage
- let Weaver populate a graphics-request database directly for P.I.G.
- let P.I.G. return completed graphics directly into Weaver QC
- pass QC-approved graphics downstream to Poetry Please automatically

## Operational Overlaps To Resolve

These are the overlapping systems currently creating confusion. Each one needs a named owner and a cleanup plan.

### 1. Frontend Hosting Overlap

Current overlap:

- Cloud Run is the real live frontend for `weaver.buttonpoetry.com`
- the live hosted shell has previously picked up frontend changes that are now visible on `weaver.buttonpoetry.com`, including the batch review pulldown
- those live frontend changes match commits in the repo's production path, so frontend shell changes should be treated as a production code release, not as an Apps Script-only update

Plan:

- treat Cloud Run as the only production frontend until an intentional migration happens
- document one explicit production deploy path for Weaver frontend changes
- record the production release rule plainly: changes in `public/index.html`, `public/app.js`, or `server.mjs` are frontend releases and must be pushed through the production repo/host path
- later evaluate a deliberate move from Cloud Run frontend hosting to another single-host setup only if it reduces complexity

### 2. Backend Execution Overlap

Current overlap:

- local Python / Node services also provide validation and popup support

Plan:

- keep Cloud Run as the live application backend
- keep the root Apps Script project for Excerpt Update Tool / sheet maintenance only
- remove hidden duplicate write paths where possible
- keep the current backend boundary explicit: Cloud Run owns interactive Weaver reads/writes; Apps Script owns separate maintenance workflows only

### 3. Source-Of-Truth Overlap

Current overlap:

- `Excerpt Tool 1.20` is the real source sheet
- `New - Quote Creation Tool Database` is a generated working view
- reviewers can still experience both as if they are primary

Plan:

- keep `Excerpt Tool 1.20` as the only durable source of truth
- treat generated tabs as queue surfaces only
- make all permanent workflow writes land on the source tab or source-backed systems
- reduce any workflow that depends on manual reconciliation between tabs

### 4. Graphics Workflow Overlap

Current overlap:

- old downstream graphics workflow uses Google Sites + Looker Studio + export sheets
- Weaver now has the start of an in-app `Graphics queue`
- P.I.G. is the intended future creation tool, and the long-term database-backed handoff is not in place yet

Plan:

- keep the old downstream tool available during transition
- build the Weaver `Graphics queue` until it can replace book counts, row browsing, and downstream handoff
- remove the manual export dependency once Weaver can write graphics requests into a Weaver-owned database for P.I.G.
- use a dedicated `PIG - Completed Graphics` return lane as the transitional writeback target for completed graphics entering QC
- define the P.I.G. read contract:
  - open graphics requests
  - stable request ids
  - excerpt / author / poem / book payload
- define the P.I.G. writeback contract:
  - completed asset link
  - request id
  - created-at metadata
  - optional production note
- clearly separate `Needs graphics` from `Created / cleanup`

### Drive Import Matching

- Weaver should own a durable book-metadata layer that includes:
  - canonical title
  - canonical author
  - book shortener
  - release catalog / season context
- near-term: keep reading book shorteners from publishing-order / catalog sources
- long-term: move book shortener authority into a Weaver-owned metadata source so import/export logic is not depending on scattered spreadsheets
- excerpt gathering follow-up:
  - for EPUB-backed books, consider poem-title lists in table-of-contents order instead of alphabetical order
  - keep title finding easy with typeahead filtering as the user types
  - keep non-EPUB / publishing-order books clearly in manual-entry mode so the UI does not imply EPUB navigation is available
- Drive import matching should prefer this order of confidence:
  1. poem title
  2. book shortener
  3. full book title
  4. author last name
- if an author has only one open book in Weaver, last-name matching can safely stand in for book confirmation
- if an author has multiple books open, Weaver should use:
  - book shortener
  - release catalog / season hints from the Drive hierarchy
  - poem-title confirmation across candidate books
- the hardest remaining case is multiple excerpts from the same author / book / poem
- for those cases, Weaver should support an explicit import matcher that shows:
  - the returned graphic
  - the candidate excerpt texts
  - a manual radio-button selection for the correct target
- near-term follow-up: keep refining the automatic importer, but preserve the manual matcher as the safety valve rather than forcing weak auto-matches through
- near-term follow-up: strengthen safe-match logic with release/catalog context when shortener matching is weak or absent

### Poetry Please Handoff

- Weaver should trigger the Poetry Please handoff automatically when a P.I.G.-returned graphic is QC-approved.
- Poetry Please should ingest approved graphics through its own import pipeline rather than relying on Weaver to upload into the Poetry Please bucket directly.
- the Poetry Please side should:
  - copy the approved graphic from Drive / remote source into its own Google Cloud Storage bucket
  - write or update the matching Firestore record
  - keep the import idempotent so retried handoffs do not create duplicate content
- follow-up: add explicit handoff state tracking so Weaver can show:
  - pending handoff
  - handed off successfully
  - handoff failed / retry needed

### 5. Catalog / Library Validation Overlap

Current overlap:

- catalog validation comes from one data path
- excerpt-library and QI status come from another
- PDF-backed books behave differently from EPUB-backed books

Plan:

- keep surfacing catalog and library signals separately in Weaver
- explicitly label PDF-backed catalog cases as a distinct review lane
- avoid pretending all validation states are equally trustworthy
- keep improving normalization rules, but preserve the distinction between title mismatch, excerpt mismatch, and missing catalog coverage

### 6. Deployment Authority Overlap

Current overlap:

- local repo changes
- Cloud Run live service
- sheet-only workflow changes from outside the repo

Plan:

- define one production checklist for any Weaver release
- require every release to answer: frontend/backend deploy target and any sheet dependency changes
- prefer this thread / workspace as the single implementation authority for Weaver changes
- document what can change safely without a frontend deploy vs what requires one

## Release Notes We Need To Preserve

### Proven Hosted Weaver Behavior

- the batch review pulldown was a frontend change
- it came from changes in `public/index.html`, `public/app.js`, and related hosted server support in `server.mjs`
- that change did become visible on `weaver.buttonpoetry.com`
- therefore the hosted Weaver shell is not Apps Script-only; it does reflect repo-based frontend releases

### Working Distinction

- Cloud Run deploys can change live backend behavior, sheet reads/writes, payload shape, and queue logic
- Cloud Run deploys can also change hosted buttons, layout, and in-app modules
- root Apps Script deploys affect the separate Excerpt Update Tool maintenance workflows, not Weaver itself

### Minimum Release Checklist

Before calling a Weaver change live, answer these three questions explicitly:

1. Did we change `public/*` or `server.mjs`?
2. Did we change root `Code.js` or another maintenance script?
3. Which production target was actually updated?

If those answers are not written down, the release should be treated as unverified.

## Non-Goals For MVP

- full visual graphic creation inside the app
- replacing every spreadsheet workflow on day one
- complex role/permission systems

## Recommendation

Build this as a standalone internal tool first, but keep the stack compatible with Poetry Please so it can be folded in later if that proves useful.

## Now / Soon / Later

### Now

- confirm exactly which excerpts from the old collection system successfully made it into Weaver’s live pipeline, and run narrow reconciliation/backfill passes for any missing books or batches
- finish the mismatched graphic pairing lane so QC-rejected mismatch records have a dedicated place in Weaver instead of disappearing into notes only
- continue the runtime DB cutover for graphics QC and Poetry Please handoff state so sheet columns are no longer the only operational truth
- improve excerpt gathering so italicized phrases can be captured intentionally instead of relying on ad hoc `*...*` workarounds
- start the EXC handoff path from Weaver to Poetry Please with stable excerpt record IDs, runtime-backed handoff storage, and a later backfill pass for already-approved excerpts

### Soon

- keep tightening Drive import matching, naming, and manual pairing polish
- add a `QC Sweep` mode in `Graphics QC` that serves one reviewable graphic at a time and auto-advances after save
- move more graphics request, completion, QC, and handoff reads onto the runtime DB
- keep improving the EPUB-backed excerpt gathering flow and metadata source-of-truth decisions
- add an explicit catalog DB sync step from the source EPUB/catalog project into Weaver's `data/formal_catalog.db`, followed by build + deploy, so catalog updates reliably reach the live app
- complete the `Correct and recreate` rework loop back through P.I.G. with cleaner state visibility
- expand safe auto-fixes in `Needs correction` for high-confidence metadata and formatting cases, while keeping ambiguous rows quarantined for human review
- define whether the graphics handoff ledger and the new EXC handoff ledger should stay parallel long-term or merge into one generalized downstream handoff model once both flows are stable

### Later

- fully replace the legacy excerpt gathering form
- sync approved excerpts from Weaver and the old tool into the canonical excerpt database
- move more spreadsheet-owned workflow state into Weaver-owned storage
- add release-date weighting and optional quick-win ranking to `QC Sweep` once release metadata is wired into the graphics QC path
- keep refining Poetry Please downstream visibility and retry/state management
- consider TOC order instead of alphabetical order for EPUB poem lists, while keeping typeahead filtering so specific titles are still easy to find

## Handoff Ledger Notes

- Weaver now needs two related downstream ledgers:
  - a graphics handoff ledger for Weaver <-> P.I.G. lifecycle state
  - an EXC handoff ledger for Weaver -> Poetry Please excerpt delivery
- In the near term, they should stay parallel:
  - graphics and EXC have different source events, payloads, and retry semantics
  - keeping them separate reduces migration risk while both contracts settle
- In a later pass, we should evaluate whether they want to converge into one generalized downstream handoff model with:
  - stable record identity
  - content type (`QI`, `EXC`, later others)
  - source system / source record id
  - handoff status
  - target system
  - retry / error metadata
- We should not merge them early just for elegance; the better test is whether the operational states and replay/backfill rules actually align.
