# Excerpt Gathering Module Plan

## Purpose

Build a Weaver-owned excerpt gathering module that first replaces the current Google Form workflow and then grows into an ebook-native excerpt capture tool.

This module should eventually:

- feed excerpts directly into Weaver
- feed accepted excerpts into the broader excerpt database
- support book excerpts, video excerpts, and quote fixes
- make excerpt capture feel like production intake work, not generic form entry

## Current North Star

The long-term win condition is:

- a user works from an ebook or manuscript view
- highlights or pastes excerpts one at a time or in bulk
- the tool attaches the correct book metadata automatically
- the excerpt enters the Weaver ecosystem immediately

## Practical Ship Order

### Phase 1: Form Parity Inside Weaver

Replicate the current Google Form behavior inside Weaver.

Required lanes:

- Add a quote from a book
- Add a quote from a video
- Fix a quote in one of the quote tools

Minimum viable behaviors:

- manual entry for the current form fields
- separate visible lanes for book excerpts, video excerpts, and quote fixes
- catalog-backed book selection for book intake
- writes into the same sheet ecosystem the current review pipeline already reads
- no workflow regression for the review queue

### Phase 2: Better Intake Ergonomics

Improve capture speed and quality without changing the underlying sheet destination yet.

Priority improvements:

- SOP guidance in-tool
- fast keyboard submission
- quote-length feedback
- book -> poem cascading selection from the catalog
- EPUB context view from intake
- EPUB poem-to-poem navigation in reading order
- `Next poem` / `Previous poem` controls while gathering from EPUB-backed books
- bulk-paste support for multiple excerpts
- clearer fix workflows

### Phase 3: Ebook-Native Capture

Move past form parity and make the tool feel purpose-built.

Target capabilities:

- load an ebook or manuscript source
- click or highlight text to capture excerpts
- preserve source metadata automatically
- queue multiple excerpts before submit
- support batch review before ingestion
- move through poems in sequence without leaving the reading flow
- optionally record full-poem reactions while gathering excerpts

### Phase 4: Database-Native Intake

Replace spreadsheet-first intake with Weaver-owned records.

Target behavior:

- excerpt gathering writes first to Weaver-owned storage
- spreadsheet becomes transitional or downstream only
- excerpt database ingestion becomes a first-class pipeline, not a reconciliation chore

## SOP Behaviors To Preserve

These rules should stay visible in the product:

- prioritize best excerpts over total count
- copy and paste from the source, do not manually retype
- keep moving past weak “maybe” excerpts
- roughly cap at 3 excerpts per poem and 30 per book
- keep many excerpts in the 10–25 word range
- support quote correction as part of the same intake system

## Current Live Shape

The current Weaver module now covers:

- Book intake
- Video intake
- Fix intake
- author/book suggestion loading
- append into `Excerpt Tool 1.20`

This is intentionally still a sheet-first bridge.

## Next Product Steps

### Next UI/Workflow Steps

- add bulk excerpt entry for book intake
- add inline formatting reminders for quote punctuation/italics handling
- add success state that encourages rapid repeated entry
- add a dedicated “gather another excerpt” flow
- confirm whether quote-fix intake should stay in this module or fully hand off to the existing correction tools
- add EPUB navigation controls so a curator can move poem-to-poem while gathering
- explore lightweight full-poem reaction capture as a stretch layer inside the same EPUB reading flow

### Next Data Steps

- define a Weaver-owned excerpt intake table
- dual-write excerpt intake to the runtime DB as a shadow path
- define excerpt database handoff rules for accepted entries
- confirm the long-term canonical source of truth for authors, books, and poem titles
- confirm what source the legacy Google Form was using for author/book suggestion lists
- use the legacy Google Form suggestion lists as the non-EPUB fallback until that canonical metadata source is settled
- use the publishing-order sheet as the near-term source for forthcoming / non-EPUB book metadata
- decide whether the long-term book metadata path should be a database scrape, an intentional database push, or both

### Next Research Steps

- inspect the current Google Form branching and exact field wording
- inspect existing fix submissions to understand real-world correction patterns
- define what metadata can be auto-attached from ebook sources

## Acceptance Criteria For Phase 1

We can call Phase 1 done when:

- a curator can complete the three existing intake jobs in Weaver
- the rows land in the live sheet correctly
- the downstream review queue sees them without manual repair
- the tool is fast enough that someone would reasonably choose it over the Google Form

## Acceptance Criteria For The Full Vision

We can call the module truly successful when:

- it fully replaces the Google Form
- it becomes the preferred capture surface for excerpt gathering
- it feeds both Weaver review and the excerpt database cleanly
- it supports direct ebook-native excerpt capture rather than only manual field entry
