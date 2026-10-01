# Weaver Video Curation Roadmap

Last reviewed: 2026-10-01

## Task Routing Resolution

This roadmap tracks the Weaver video-curation pilot only. The live, read-only set validation below is complete, and video pilot work has resumed in the Weaver workstream.

- Excerpt Database reliability, complete-dataset lookup, and consumer exports belong to the Excerpt Database workstream.
- Firestore deployment and separation from the legacy excerpt sheet belong to the Weaver workstream.
- Authenticated reviewer persistence and Poetry Please video handoff checks belong to the Weaver video-curation workstream.
- Do not infer that an excerpt is absent from the canonical database by searching only a partial Google Sheet export.

## Current Status

Engineering completed for the priority pilot:

- The five current 2026 Drive sets can be represented with Drive file IDs as stable identities.
- Reviewer ratings persist independently from excerpt selection.
- Full priority rankings retain every video and ignore blank legacy scores.
- Replacement Drive files can recover reviews only through a unique exact filename match.
- Reviewer-specific progress export includes reviewed, remaining, rating, excerpt count, and review time.
- Poetry Please handoff rejects raw footage and publication-restricted sources.
- Poetry Please receipt is not considered successful unless canonical video identity, reviews, and selected excerpt links are confirmed.
- The focused video handoff suite passes all 12 tests as of 2026-10-01. The broader JavaScript suite was not rerun for the FV/OM change.

## Immediate: Priority Video-Set Pilot

Next pilot gates:

1. Completed 2026-09-27: all five live sets load in Weaver: BPL Charm City, MN Writers Respond @ The Loft, MPMU Finals, Ollie Schminkey Action Cam, and Publisher's Poetry Slam 2026 - Camera Y.
2. Run one reviewer through an end-to-end progress test and compare the progress export with the visible Weaver state.
3. Verify `dislike`, `meh`, `like`, and `moved_me` all persist after reload without requiring excerpt selection.
4. Approve one safe test video with selected excerpt IDs and a publishable asset.
5. Confirm Poetry Please returns canonical video identity and acknowledges every review and selected excerpt link.
6. Record matched, missing, duplicate/version, and unresolved counts for the four-set pilot.
7. Only after those checks, decide whether the priority pilot is ready for normal staff use.

### Live Validation Snapshot: 2026-09-27

- BPL Charm City 2026: 54 videos; no blank author/title fields; every video has at least one linked Weaver excerpt.
- MN Writers Respond @ The Loft: 12 videos; no blank author/title fields; every video has at least one linked Weaver excerpt.
- MPMU 2026 Finals: 16 videos; one video has no linked excerpt: Keaghan O'Brien / `Tilt (slash) Shift` (`1vW708q89C8fKxxi2S5Gi2oOBTXtFOQG5`).
- Ollie Schminkey Action Cam: 11 videos; no blank author/title fields; every video has at least one linked Weaver excerpt.
- Publisher's Poetry Slam 2026 - Camera Y: 22 videos; Logan Lopez / `Audition` has no linked excerpt (`1vZqbwF4sJS4SXOHZZw3Gb2PTaxpa1Eem`).
- Sophie Wang (`1t4PYbVXgdDU9jPURqbeY7W9s6MglDJBj`) is marked `opted_out`, has no poem title, and is correctly publication-restricted.
- Total: 115 videos, two without linked excerpts, and one correctly restricted opt-out.
- Aggregate progress and rating-persistence verification remain pending because those operations require the approved signed-in Weaver admin/reviewer workflow.

Operational rules:

- Treat the Drive file ID as stable source-video identity and retain exact folder/file provenance.
- Keep raw footage links and reviewer identities in Weaver.
- Send only approved excerpts, publishable media, and appropriate curation fields to Poetry Please.
- Enrich from Footage Inventory when available without blocking initial review.
- Deployed in `weaver-00448-76v`: a score of 9.0 with five distinct normalized reviewer emails triggers auto-advance, but that automatic trigger has not been live-tested. A controlled manual same-file FV handoff passed. Source and final-asset classification use Drive folder ancestry; an FV may be its own final asset, while OM requires a separately verified FV. Firestore conditionally claims handoffs once; uncertain responses require reconciliation rather than blind retry. Later reviews remain allowed without a second initial send. Revisit score formula and later-review updates after the pilot.

### FV / OM Intake Contract

- `FV` means Finished Video; `OM` means Original Media. Record the type when the footage first enters Weaver, keyed to the exact Drive file ID and stable video/poem identity. Do not infer type from filename, resolution, review history, or whether the handoff URL differs from the review URL.
- Track release clearance separately from FV/OM. An FV is eligible for direct Poetry Please handoff only after release clearance and curation approval. An OM can be reviewed and excerpted in Weaver but must go to editing; attach its later FV without changing the review/video identity.
- Unknown or mixed intake remains unclassified and blocked from automatic publication until explicitly classified. Preserve an audit trail for classification changes. Confirm the pilot folders' actual FV/OM types before backfill.
- The exit-time different-file check has been replaced with FV, release, and asset-identity validation. The exact reviewed file may be the publishable FV; a different file is not automatically publishable. Poetry Please accepted one controlled same-file FV handoff on 2026-10-01.
- Auto-advance currently runs when a review is saved; it does not retroactively process all previously qualified videos. The manual FV canary passed, but test the automatic trigger and an OM-to-edited-FV transition before bulk reconciliation.

### Review FV / OM Classification After Pilot

- The first implementation classifies video sources from their Drive folder ancestry at review intake. The four edited pilot folders resolve to FV; the Camera Y pilot resolves through the Original Media structure. Unknown or mixed paths do not auto-publish.
- Review this design after an OM-to-edited-FV handoff. The controlled FV same-file handoff passed; still verify release clearance independently and confirm reviews/excerpts retain their original video identity when an OM gains an FV.
- Likely improvements: use a maintained folder-ID registry or Footage Inventory role evidence in addition to folder-name markers; show classification evidence and ambiguity to admins; persist and audit deliberate reclassification; handle moves, shortcuts, and assets recorded in both FV and OM roles. Do not turn a folder label alone into permanent content truth.
- Deployed 2026-10-01 as `weaver-00448-76v` (100% traffic). Jay Ward / "Critical Blues Theory 101" (`1XR1vNwCYvcj5alxx3Qri0deH9IqM4UFK`) passed the same-file FV canary: Weaver sent one handoff, Poetry Please created `WEAVER-VV-58A0A3F8AA4BA88E2BC3D599`, its admin row shows 11 reviews and 6 selected excerpt IDs, and its video plays from Poetry Please storage. A redundant final-file ancestry lookup initially returned unknown; the deployed fix reuses the already verified FV classification when the exact reviewed Drive file is the final asset. The sent view contains this record once.

### Priority Video And Video Lanes

- Locally implemented: group video sets into Priority Video and Video using their existing records; this changes only the displayed lane, preserving reviews, excerpts, scores, source identities, and publication gates. Deployment and live verification remain pending.
- Confirmed transition rule: move an event out of Priority Video only when every video in the event has at least five reviews from distinct normalized reviewer emails and at least three saved excerpts. A video below either minimum keeps the whole event in Priority Video. This is a curation-progress rule, not an authentication gate; the user confirmed existing tracked emails are sufficient.
- Revisit the five-review requirement before volunteer curators affect lane completion.

### Live Deployment And Canary: 2026-10-01

- The approved `deploy-cloud-run.sh` workflow deployed Weaver revision `weaver-00446-jrg` to 100% of Cloud Run traffic. The live UI reported that revision and loaded the video scoreboard and review queue.
- At the time of revision `weaver-00446-jrg`, the Poetry Please send canary had not passed. The earlier Black Chakra / "He's Too Loud, He Sounds So Urban" handoff was marked `failed`; do not retry it until its Poetry Please outcome is reconciled. The later Jay Ward same-file FV canary passed in `weaver-00448-76v` as recorded above.
- Next: test the automatic five-review/9.0 trigger on one controlled FV and test an OM-to-edited-FV transition. Reconcile uncertain historical handoffs before retrying them; do not bulk send from the scoreboard yet.

### Video Excerpt Review Lane

- Open decision: when a video depicts a poem already represented by book-sourced excerpts, should matching excerpt text resolve to one canonical excerpt while retaining both original intake records and the video relationship? Do not assume that separate book and video review lanes, the excerpt DB, or Poetry Please merge these automatically. Define matching and provenance rules before cross-source deduplication; the `approved / 3` target should ultimately count distinct canonical excerpts, not duplicate source rows.
- Deployed in `weaver-00449-jxp`: the `Video excerpts` review picker groups pending excerpts by event and source video. It shows a provisional `approved / 3` count based on distinct normalized video-excerpt text and refreshes after decisions save. The existing approval chain and real book metadata are unchanged. A direct endpoint smoke check was unavailable from the local sandbox because DNS resolution failed; Cloud Run reports 100% traffic on the revision.
- The former combined pseudo-book had 150 pending rows, 113 distinct displayed excerpts, and 28 additional pulls. Complete the event/video review experience by making below-target videos easy to find even when they have no pending excerpts; do not rewrite real book metadata to the event name.
- Show `approved / 3` for each canonical video, counting distinct approved excerpts rather than gathered pulls. Make below-target videos easy to find; update counts after decisions save.
- Rank candidates within a poem/video by the number of distinct reviewers who selected the same normalized excerpt. Preserve each original pull and reviewer attribution; ranking is guidance, not automatic approval. Define a stable excerpt grouping key before merging near-duplicates.
- Keep approved video and excerpt handoffs separate while retaining their shared video/poem source identity. Do not add another excerpt-content source of truth.

## Later: Historical Curation and Video Backfill

- Normalize the 2018-2026 quarterly curation responses into reviewer-level durable review records.
- Reconcile those records with the historical Curation Database `All_Responses` archive and `Scoreboard` aggregates.
- Join exact video assets through Footage Inventory identity and use the existing Looker export as a delivery and verification artifact, not source truth.
- Preserve historical 0.0-10.0 values and derive categorical compatibility without overwriting the original score.
- Backfill approved video/poem curation aggregates into Poetry Please without sending unedited footage links or reviewer PII.
- Produce matched, ambiguous, duplicate/version, missing, and unresolved reconciliation counts before retiring old operational forms.

## Related Work That Must Remain Tracked

### Excerpt Database Reliability

- Database-wide presence checks must query the canonical database or the complete export, never an arbitrary Sheet row range.
- Regression case: Rudy Francisco / *Helium* / “Tragedy and silence ... exact same address” resolves to “Complainers”; it was previously missed because only the first 15,000 rows of a 47,461-record export were searched.
- Give QI and video consumers a stable full-dataset lookup contract with normalized quote matching.
- Preserve source IDs and exact provenance when approved video excerpts enter the excerpt database.

### Weaver And Legacy Sheet Separation

- The previously prepared Firestore excerpt dual-write work is recoverable from commit `4764253`, including `WEAVER_EXCERPT_FIRESTORE_CUTOVER.md`, `backfill_weaver_excerpt_firestore.py`, and `weaver_runtime_sync.py`; reconcile it with current Weaver before restoring or deploying it.
- Do not freeze the legacy excerpt sheet until Weaver intake, review, corrections, and approved export no longer depend on it.
- Require a sheet-disabled end-to-end test before declaring the legacy sheet archival-only.

### Catalog Reconciliation

- Continue edition and title normalization decisions in `CATALOG_RECONCILIATION_ROADMAP.md`.
- Keep catalog ambiguity separate from missing excerpt coverage; neither should be inferred from a partial consumer export.
