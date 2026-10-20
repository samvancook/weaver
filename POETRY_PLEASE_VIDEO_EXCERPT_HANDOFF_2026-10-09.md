# Poetry Please handoff: book-optional video excerpts

Date: 2026-10-09
Status: Historical checkpoint. Verify current Poetry Please behavior before retrying these rows.

## Live canary update (2026-10-09)

- Poetry Please PR #165 deployed its bookless video-linked EXC path. Weaver deployed its matching handoff and one-record admin retry as `weaver-00459-dvj`, serving 100% traffic.
- Only row 10981, source ID `weaver:477591ee-f604-4a0b-8125-09a976d9101b`, was retried. The live UI reported `EXC handoff failed: duplicate_excerpt_text`. Weaver did not mark it sent. Rows 11318 and 11320 were not retried.
- The live Poetry Please write guard rejected another EXC with the same normalized excerpt text, created a duplicate-review record, and returned `409 duplicate_excerpt_text`. The Weaver import result did not supply the existing canonical excerpt ID as a successful link receipt.
- Poetry Please should preserve the existing canonical excerpt and record the video/source occurrence against it, or return an explicit reviewable duplicate outcome with the canonical excerpt ID and video association. Do not create a second canonical EXC or infer a book. Weaver needs a machine-readable receipt before it can mark the handoff complete.

## Original failure (before PR #165)

Three approved BPL Charm City 2026 video excerpts failed the Weaver-to-Poetry-Please handoff with `Missing required publishing metadata: bookShortener, releaseCatalog`. They have source-video links and an event, but no captured book attribution. This does not prove the poems are not in books; do not guess book data or use the event as a book title.

- Weaver source IDs: `weaver:477591ee-f604-4a0b-8125-09a976d9101b` ("I Didn't Read The Love Poem"), `weaver:e9f23eff-c4c5-49f8-b779-0ce4eeb38cc0` and `weaver:0abdf5b2-1ee2-4b6c-acd1-6d8681316d92` ("Copaganda").
- Before PR #165, Poetry Please required a book-derived excerpt ID and a nonblank book. PR #165 changed that bookless video-linked path.

## Bookless contract delivered by PR #165

Accept an `EXC` from Weaver when it has a stable `sourceRecordId`, author, poem title, excerpt text, and a verified source-video identity (`sourceVideoFileId` / `sourceVideoRecordId`, preferably `sourceContentId` for the imported canonical video). Book, `bookShortener`, and `releaseCatalog` remain optional **only for this video-linked path**. Keep the existing book-backed excerpt path unchanged.

Use the stable Weaver `sourceRecordId` for idempotent identity or another deterministic, collision-resistant video-excerpt ID. Do not generate a fake book shortener or collapse distinct pulls solely because poem titles match. Preserve the video link, event, approval provenance, and any real book metadata supplied later. Return the created/updated ID; return an explicit error if the video identity cannot be resolved rather than silently creating an orphaned excerpt.

## Canary and acceptance

1. Reconcile row 10981 against its existing canonical excerpt, without reapproving or creating a new source record.
2. Confirm Poetry Please creates or updates one EXC linked to the intended canonical video, with correct author, poem, text, event, and no fabricated book/catalog.
3. Repeat the handoff and confirm no second excerpt is created.
4. Verify a book-backed EXC still imports as before.
5. Only then consider the other failed BPL records and reconcile Weaver status with Poetry Please receipts.

Approved video and excerpt handoffs remain separate while retaining their shared source identity.
