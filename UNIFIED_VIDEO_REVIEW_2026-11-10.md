# Unified Weaver video review — November 10, 2026

Use https://weaver.buttonpoetry.com/ as the canonical tool. The existing footage-bridge and video-queues URLs serve the same verified revision, `weaver-unified-review-f1ccc5bff2`, so Footage Helper bookmarks continue to work. Main traffic is 100% on that revision; generation and observedGeneration both read 582, Ready=True.

The release combines the newer main Videos workspace and identity/receiver features with the per-video personal queue, saved-review history, library and bounded save behavior. Source is based on main repair commit 019588d, incorporating queue commit 16efb6a selectively. Existing backend review, excerpt, gate, receipt and inventory writers are unchanged. Main backend changes are read-only library/progress projections.

Event labels now state “N awaiting your review”, “N reviewed by you” or “N videos total”. Missing email shows an identity prompt instead of claiming personal progress. Both rating controls retain optional excerpts, a 30-second video-write limit, one read-only outcome check for uncertain rating writes, preserved entries and no automatic write replay. Reviewer/event/rating controls and navigation stay locked while saving. A failed shared-progress refresh reports the saved rating separately.

75 focused checks passed: queue/library filtering, counts and labels, exact task links, rating/timeout behavior, identity/receiver handling, excerpt delivery, authentication, Helper receiver integration and the main Videos workspace. Live canary and all three production URLs matched all three public source checksums and exposed all 122 current videos. The image also includes the exact tested server and handoff helper, as recorded in the immutable build manifest.

Sam's exact Tilt/Shift review was read from the canonical saved records: one record, rating moved_me, updated 2026-11-10T15:25:09Z, with the user's notes intact. Live main UI and APIs show five reviewers, three excerpts and reviewComplete=true. All sixteen MPMU videos meet the review minimum; zero await Sam in the priority queue, one is in Sam's reviewed history and all sixteen remain in the library. This does not constitute publication approval, an actual Poetry Please send, or inventory event closure.

Complete pre/post-deployment snapshots of 981 curationReviews, 43 videoCurationGates and three footageEventIntakes compared exactly equal. No production rating, excerpt, gate, media or inventory edits were made by this release. The real review was submitted by Sam before deployment; save behavior after this consolidation was exercised in isolated tests, not through a fabricated production review.

Private evidence: /private/tmp/weaver-unified-review-2026-11-10/ (manifest, immutable source archive/build results, traffic requests/readbacks, canary/live readbacks, full before/after records and screenshot). These records may contain private infrastructure configuration and must not be committed or broadly shared. No rollback or undo was performed.

Human guide: [Reviewing videos](VIDEO_REVIEW_QUEUE_GUIDE.md).
