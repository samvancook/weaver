# Durable video-excerpt delivery

Approving a video excerpt saves the editorial decision first and queues delivery in Weaver's existing Firestore `excerptHandoffs` ledger. The save response distinguishes saved approvals from delivery problems. It does not wait for Poetry Please to import the excerpts. Book and full-poem handoffs retain their existing path.

The existing authenticated video-pipeline Scheduler worker runs the queue. No additional production scheduler, collection, or IAM grant is required. The worker verifies current source approval and an identified parent video before taking a per-source, conditional-write claim. It rechecks the approval after claiming, then makes one automatic delivery attempt. A repeated approval or overlapping worker cannot reset that attempt.

## Delivery states

- `queued` / `waiting_video`: no request attempted; eligible for a claim once the parent video is identified.
- `sending`: an active, leased attempt. Other workers leave it alone.
- `acknowledged`: PP returned an exact canonical excerpt ID; independent receipt verification remains due.
- `received`: the receiver receipt was verified and reconciled into Weaver.
- `needs_reconciliation`: an attempt failed, expired, or had an uncertain outcome. The worker only reads receiver receipts; it does not resend.
- `needs_source_review`: the approval was withdrawn or its saved source version changed. It requires explicit resolution.

The Videos delivery panel displays these ledger states. A successful editorial save is not evidence of successful PP receipt. The legacy source Sheet's delivery columns are not updated by the asynchronous worker; the durable ledger and receipt check are authoritative for these deliveries.

Existing delivered records remain intact and are not adopted into the new queue. Existing uncertain history requires reconciliation, even if someone saves the approval again. Generic handoff upserts and manual retry controls cannot bypass managed delivery claims.

## Failure and recovery

If approval saves but queue persistence fails, Weaver reports that distinction. Refresh the authoritative source and delivery ledger before further action. Do not blindly resubmit a batch or reset an attempt. If PP accepted a request but the response was lost, an exact-source receipt must settle the outcome. A bounded recent-import listing cannot prove absence; the item stays unresolved until positive receipt evidence or a separately authorized repair establishes the next action.

A stored source-row number is not used to write a later delivery result back into the Sheet: rows can move. Changing or withdrawing an editorial decision never deletes a receiver record automatically.

## Readiness boundaries

Video selection still uses the approved 8.5 score and five independent reviewers, with excerpt counts visible and no ranking bonus. FV/OM resolution, asset linking, release decisions, excerpt approval, verified receiver media, caption selection and Calendar readiness remain separate gates. An excerpt receipt does not make a video Calendar-ready, and unknown media identity never authorizes an editing request.

## Verification

The deterministic tests cover repeated approvals, conditional claims, overlapping workers, changed and withdrawn sources, uncertain outcomes, expired leases, legacy history preservation and a successful source save followed by queue failure. A deployment must also pass an authenticated, read-only preview against the existing Firestore backend. Do not manufacture an editorial approval as a deployment test; record the first real managed delivery's exact receiver receipt when it occurs.
