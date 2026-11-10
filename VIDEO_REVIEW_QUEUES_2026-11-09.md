# Video review queues and library — 2026-11-09

## Released scope

The existing footage workflow's Weaver URL now serves revision `weaver-video-queues-51384aa452`:

https://footage-bridge---weaver-iasomqy6ja-uc.a.run.app/?videoSet=mpmu-2026-finals

This release changes queue presentation and adds a read-only video library. Each video can leave Priority independently when it meets the existing minimum of five distinct normalized reviewer emails and three distinct saved excerpt IDs. Existing review upserts, publication decisions, release rules and handoff writers are unchanged.

The three views are **My review queue**, **My reviewed videos** and **Video library**. All support event selection and author/title/event/filename search. The library also filters by Needs review, Review complete, Ready for handoff, Needs editing, Needs a check and Sent to Poetry Please. Explicit Vertical Version A/B filename labels distinguish separate assets; reviews remain attached to their existing set/file identity.

A personal queue hides videos already rated by that reviewer and videos whose shared review is complete. Reviewed videos remain accessible with their saved rating and notes. Revisions to a rating use the same existing reviewer/file upsert, rather than adding another vote. A failed follow-up read is reported as a refresh failure after the successful save; the interface keeps the saved rating. Individual items carry their own event identity when browsing All events. Reviewer changes invalidate the loaded personal view, stale loads are ignored, and a changed selection cannot overwrite another video's local review.

## Observed live result

The final bridge API and interface preserve all **122 videos** in the six current sets. **Nine** still need Priority review: four BPL videos, one MPMU video and four PAGES videos. MPMU retains all sixteen library entries while its personal review queue offers only **Keaghan O'Brien — Tilt (slash) Shift** (four reviewers, three excerpts). One more reviewer meets its existing minimum.

Sam's existing **Black Chakra — Click** review appears in My reviewed videos with its saved Like rating. The MPMU library visibly distinguishes Vertical Version A/B and keeps videos with existing Poetry Please handoffs available. No real rating was submitted in this verification.

The release read model currently classifies 76 files as Review complete, 9 Needs review, 19 Sent to Poetry Please, 7 Needs a check and 11 Needs editing. No file is presently classified Ready for handoff. Sent is a saved accepted-handoff status, including an unambiguous prior media identity; it is not a fresh receiver, durable-copy, public-release or calendar-placement certification. Ready for handoff also requires an eligible score or existing approval, FV, release clearance and a confirmed media identity. Ambiguous historical associations remain Needs a check.

A BPL item already has eleven reviewers but only two saved excerpts. Existing reviewers can use My reviewed videos to add the missing excerpt without another independent vote. The minimum was retained, not changed or newly negotiated in this release.

## Verification and boundaries

- 74 focused tests pass across queue behavior, saved-review retrieval, reviewer/file uniqueness, historical accepted-handoff identity, failed-read/save races, task links, inventory hydration/provenance, video handoff contracts and excerpt release handling.
- Live browser checks verified the one-item MPMU queue, sixteen-item library, distinct A/B versions and Sam's saved-review history. API readback verifies the served app checksum and all current files. Save/disappear and update behavior were exercised in isolated fixtures; the next real staff review still needs to exercise the full production save loop.
- Exact before/after Firestore documents compare unchanged: **980 curation reviews, 43 video curation gates, three footage intake receipts**. This change and its verification issued no score, gate, receipt, inventory sheet, closure or media mutations.
- All three Helper connections still return transferComplete: sixteen MPMU entries and the two four-entry PAGES intakes. Original PAGES and MPMU remain closed; the separate vertical PAGES intake remains open. Anonymous receipt reads remain refused; the authorized Helper path passed.
- The wider priority-dashboard suite has six passing checks and one pre-existing assertion failure. That assertion expects current candidate excerpt IDs to replace a saved gate's selected IDs before handoff, while the existing writer preserves the gate selection. The same mismatch exists in the pre-change source. This release leaves that writer intact; the intended selection policy requires separate reconciliation before claiming that broader publication check is resolved.

## Deployment evidence

The base image was pinned to the exact existing footage bridge digest. Its server, handoff module and three public files were independently downloaded and matched to the pre-change Git HEAD. The release image inherits that immutable runtime/data image and overlays only those five files. No local catalog/database snapshot or unrelated main-site code was deployed.

A canary was verified before updating the existing `footage-bridge` tag. Cloud Run generation/observedGeneration match with Ready=True. The main service remains pinned at 100% to `weaver-00535-4bn`, preserving its newer unrelated work. The prior bridge revision remains retained; no rollback was invoked. Moving this slice into the current main-site branch/release remains a separate integration step.

Owner-only baseline source, manifests, builds, requests, API responses and exact data snapshots are retained under `.footage-recovery/video-queue-release-2026-11-09/`, excluded from Git. The final iteration is `final-ui/`. The public screenshot is `/private/tmp/weaver-video-library-2026-11-09.jpg`.

## Next practical steps

1. Have a real curator finish Keaghan's fifth review and verify that it leaves their queue and Priority, while remaining in the library. Do not invent a vote to force the transition.
2. Bring the next verified inventory event through the existing Helper connection and task-link return, using these same views. Continue preserving old curation destinations and historical reviews.
3. Resolve the seven Needs a check identities/clearance cases and the handoff-selection assertion in a focused publication pass. Verify one actual authorized FV handoff and receiver receipt before expanding that path; sent videos remain available without another initial send.
4. Merge the queue slice into the current main Weaver release while preserving its newer changes. Add durable Helper return-job continuation so owners need not keep a page open. Neither step requires reworking footage intake or closing events automatically.
