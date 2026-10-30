# Footage Inventory → Weaver bridge

## Progress return and retry reliability — 2026-10-30

**Check Weaver receipt** reads current Weaver reviews and decisions by exact curation event/file IDs. Helper shows the latest review count per person, linked excerpt count, editing/held/ready state, and the last saved Poetry Please acknowledgement (including canonical video link and media-copy state when present). It reports the check time. An unreadable ledger is shown as unavailable, never as zero work. Calendar placement remains explicitly not checked; a saved PP receipt does not prove scheduling or publication. Historical reviews on another Drive ID are not silently transferred.

Weaver playlist/candidate context retains inventory event ID/name, response row, spreadsheet ID, local asset family, OM source and FV deliverable. Curation decisions preserve that provenance across later edits. This metadata is not a canonical poem/text identity, and downstream PP provenance persistence is a separate pending receiver change.

Same-manifest retries cannot remove previously accepted file relationships or downgrade/change a confirmed OM/FV role. The original receipt is retained on a failed retry, and the user can check it again. Receipt readback verifies the actual saved record, not just its manifest digest. Missing files and unavailable target queues are named in Helper. These checks do not add approval steps or write to the source inventory.

## OM/FV routing — 2026-10-30

Media role and orientation are separate fields: `mediaType=OM|FV|unknown` and `format=HV|VV|square|unknown`.
The reviewed PAGES correspondence explicitly classifies its four inventory files as OM and its four existing Weaver files as FV, based on the user's confirmation and visual branding/credit checks. Orientation alone never assigns a media role.

Each accepted title carries `workflow.originalMediaFileId`, `finishedVideoFileId`, `publicationAssetFileId`, `editingStatus`, and `publicationStatus`. A linked FV sets `finished_video_linked`, preventing this connection from implying a fresh editing job. OM without FV is `needs_editing`; unidentified media remains `media_type_needs_review`. These are routing descriptors, not writes to an editing queue. Existing Weaver curation decisions still control actual editing and Poetry Please handoffs.

Only the FV is selected as the publication asset. The existing Poetry Please sender still requires actual curation approval and a verified final FV. No scores, excerpts, identity approvals or publication clearance are fabricated from a successful Helper receipt. Calendar integration must refer to the canonical video received by Poetry Please; no calendar posts are created by this bridge.

PAGES currently has zero saved video reviews, video-linked excerpts or curation gates (authoritative ledger checked 2026-10-30). Its four video relationships are saved, but a live Poetry Please transfer remains pending actual curation work. Existing book excerpts by Pages Matam are separate and are not treated as video curation.

The bridge reads current open-event titles from Helper, previews their Weaver connection, then saves a receipt when a person chooses Link / Send. It never writes to Footage Inventory sheets, moves media, publishes to Poetry Please, or closes events.

Source Helper: https://script.google.com/home/projects/1aGmHhfSQujviWDRN7WY5mMJEVbG3t81O0MucZ0JWLN266vbvoNoWa0k1/edit

## Current update: existing events and video versions

- Existing curation event IDs and reviewed source file IDs are retained. An association receipt does not create another queue or rewrite reviews/excerpts.
- PAGES Matam footage supplied by Rudy is the first reviewed event correspondence. Four landscape source files connect to four vertical files in `pages-matam-san-diego-2026-labc`. The mapping is in `footage_event_links.mjs` and the original and edited IDs stay distinct. It is based on the named source/FV event folders and the reviewed four-title crosswalk, not byte identity or a frame comparison.
- For other events, exact existing Drive IDs can identify a single existing queue. Partial/ambiguous matches are surfaced rather than silently duplicated. Titles alone never merge events, recordings, or canonical texts.
- Each accepted row records related assets with independent IDs, measured `HV`/`VV` format, roles, and a stable family ID. Multiple inventory versions can connect to the same existing curation asset without adding review tasks. The existing curation version remains the review target; related originals are reference links.
- `textReference.canonicalTextId` stays null and `canonicalTextLinkStatus=not_linked` until the authoritative underlying text is linked. These local asset families are not fabricated Poetry Please text IDs. Existing reviews are never pooled across versions or recordings.
- The inventory's exact author/feature/workshop contract status is accepted for internal curation, along with `Have video release`. Publication restrictions are retained. This does not assert legal clearance or infer an absent release from the author-contract status.
- `transferComplete` confirms only the submitted association/transfer and current metadata accessibility. `eventCompleteness=not_assessed` and `readyForClosureReview=false` are always returned. Adam's known incomplete intake cannot become closure-ready because two listed videos were received.

## Staff steps for the separate live test

1. Open the test Helper and select **PAGES footage from Rudy Francisco 2026**.
2. If needed, choose **Connect Weaver (owner)** once.
3. Choose **Preview Weaver connection**. Expect the existing San Diego event, four accepted title rows, and an HV/VV pair for each. Preview writes no receipt or queue.
4. Choose **Link to existing Weaver event**.
5. Choose **Check Weaver receipt**. Success means the four submitted relationships are recorded; the event remains open. The existing Weaver playlist should still have four curation items, with related HV/VV links visible.

Reviewers keep using the existing Video curation interface. To review a different version as its own review target would require an explicit curation choice; these source-reference links do not silently change which asset a rating belongs to.

## Downstream boundary

Current Poetry Please `main` was checked during this update. Its Weaver video receiver still explicitly accepts only `VV` (`functions/weaver-video-intake.js`, `buildWeaverVideoIntake`). This change does not alter that receiver or claim a completed HV import or canonical-text association in Poetry Please. That remains downstream work, owned by the Weaver → Poetry Please workstream. No publication or downstream send is part of this pilot.

## Runtime and deployment

Weaver uses the existing `footageEventIntakes` collection in its Firestore ledger. Same-manifest retries keep one receipt; changed manifests or changed queue/asset routing return a conflict and preserve the old association. Both source and curation assets are rechecked, as is existing queue membership. Existing-event receipts never shadow the static queue during restart hydration.

Configuration remains:

- `WEAVER_FOOTAGE_BRIDGE_ENABLED=true`
- `WEAVER_HELPER_TOKEN_AUDIENCE` = exact signed Helper OAuth audience
- `WEAVER_HELPER_ALLOWED_EMAILS` = verified deploying owner's email
- Helper Script Properties `WEAVER_BRIDGE_URL`, `WEAVER_BRIDGE_ENABLED`

The reviewed test endpoint is fixed in `ConnectionSetup.gs`. Signed identity tokens stay server-side; only audience/email metadata is exposed to the owner. Google documentation: https://developers.google.com/apps-script/reference/script/script-app#getIdentityToken().

Apply the small `Code-doGet.patch`, `IntakeQC.html`, `WeaverBridge.gs` and `ConnectionSetup.gs` to a fresh current-source pull. Preserve all other source files. The owner-only `?page=weaver-preview&event=...` route runs the same read-only preview for deployment verification, without saving a receipt. A preview may use the fixed reviewed endpoint without storing connection configuration. Sending still requires the explicit connection flag.

`appsscript.live-test.json` preserves DOMAIN / USER_DEPLOYING and makes Sheets/Forms, read-only Drive, email/OpenID and external-request scopes explicit. Never push the entire Weaver repository through clasp. Read back every affected source file and deployment version. Keep normal Helper v43 and production Weaver traffic unchanged during the separate test rollout.

## Validation and limits

44 selected JavaScript checks and 12 Python checks passed, including existing video handoff/claim flows. New tests cover existing-queue preservation, author-contract curation, HV/VV identity separation, read-only preview, retry/hydration, changed-routing protection, loss of either asset/queue, ambiguous matches, and multiple source versions in one family. External services are mocked in these tests; deployment readbacks and the actual PAGES preview are separate evidence.

The wider dashboard suite previously had one source-pattern assertion that already failed on unmodified base; this update does not claim that suite passes. No frame/audio identity verification or curator playback has been performed. A successful transfer is not proof of a complete event or cleared publication rights.
