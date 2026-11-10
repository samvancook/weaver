# Main Weaver video submit repair — 2026-11-09

## Evidence and scope

Sam reported submitting Tilt/Shift on https://weaver.buttonpoetry.com/ and supplied a screenshot showing “Add at least one excerpt before submitting this legacy video intake.” The canonical current FV is `1vW708q89C8fKxxi2S5Gi2oOBTXtFOQG5` in `mpmu-2026-finals`. Read-only checks found no saved Sam review for that file, four distinct reviewers and three excerpts. Sam's only saved canonical review was the earlier Black Chakra / Click review.

The deployed main interface selected an event when **Review videos** was clicked but did not actually load its Weaver playlist. It could therefore leave stale or legacy context active. This explains a route into the legacy validation path; the exact state of Sam's original browser was not captured. The other reported spinner was not reproduced against a real save. The save request had no bounded wait.

## Changes

- Review videos now loads the selected Weaver event and clears the legacy-form URL, without waiting for the unrelated book catalog.
- Event/reviewer changes reload the task; stale responses cannot replace a newer selection. Payloads use the actual loaded item's event/file identity.
- Both save buttons support a rating with zero excerpts for a loaded Weaver video. A legacy form with no quotes gets an actionable loading instruction, with its entry retained. Manual legacy excerpt intake remains available.
- Video save requests stop waiting after 30 seconds. A rating save with an uncertain outcome gets one bounded, read-only check of that exact reviewer/file/rating/notes and a fresh save timestamp. No write is replayed. If it still cannot be confirmed, the entry remains and the interface says the result is unknown.
- Pending controls always recover. A confirmed rating followed by an excerpt failure is reported as a saved rating plus an excerpt problem, rather than claiming everything failed.
- Main-site task links support `videoSet` and an exact `video` ID. A task already reviewed or unavailable never falls through to a different video.

## Verification and deployment

52 focused checks pass, covering the loaded-event route, optional excerpts, retained drafts, bounded timeout, exact readback, no mutation replay, partial success, direct task targeting and preserved legacy entry, plus existing handoff/identity/authentication tests. No real rating was submitted in these tests.

The repair is an isolated worktree based on main-source commit `7c4f053`. It overlays only `public/app.js` and `public/index.html` on the immutable existing main image (`weaver-00535-4bn`). The server, backend writers, publication logic, data, credentials and permissions are unchanged. Other current tags, including the footage bridge queue release, remain intact.

The canary passed exact served-source checksum checks and loaded Tilt/Shift from the correct MPMU set. The deployed main revision is `weaver-main-review-fix-f99600aeea`. Cloud Run readback has matching generation/observedGeneration, Ready=True and 100% main traffic on that repair. Main-site source checksums match the tested build. The existing main video workspace remains available.

A diagnostic empty POST to the live save endpoint was rejected by automatic approval review before execution because it might cause a write. The investigation used read-only production checks and isolated save tests instead. No production review was resubmitted, fabricated or changed. The user's real save remains to be verified after retry; deployment is not proof that it saved.

Owner-only deployment/source manifests and readbacks are retained in `/private/tmp/weaver-main-submit-fix-2026-11-09/`, with the final rollout under `release/`. The earlier footage-bridge queue slice remains a separate release; this repair does not replace the main workspace with that older branch.

## Direct retry task

https://weaver.buttonpoetry.com/?videoSet=mpmu-2026-finals&video=1vW708q89C8fKxxi2S5Gi2oOBTXtFOQG5&email=sam%40buttonpoetry.com

Open the task, select the intended rating and click **Save Rating & Next**. Excerpts are optional here. Confirm the canonical Sam review and fifth distinct reviewer before declaring the real save-and-completion loop complete.
