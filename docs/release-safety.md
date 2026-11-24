# Weaver release safety

Weaver already has its own repository. Keep video, book intake and graphics
review together while they share authentication, source records and the runtime.
PP, Catalog, PIG and Calendar continue to own their own implementations.

Before the existing exact-commit release route, run:

```sh
node scripts/check-release-source.mjs FULL_APPROVED_COMMIT > /private/tmp/weaver-source-preflight.json
```

This read-only guard fetches current main, requires a clean checkout at the full
approved SHA containing main, checks Docker/source-upload dependencies, and
records runtime-source hashes. A newer main means reconcile and test first.
Keep receipts outside Git. This does not change credentials, build, traffic or
the established release route. Use exact approved source, not an unpinned branch.

For an immutable-image patch release, inspect the actual serving image and
preserve every newer module before building the patch. The source preflight
does not establish that a patch image contains all current modules.

Before promotion, re-read the current serving revision and main. If either
changed since preparation, stop and reconcile. After promotion, read back the
image/module manifest, main and operational aliases, then the specific workflow
being repaired. Health alone does not prove delivery or saved decisions.

## One shared release claim

`scripts/release-lock.mjs` uses `weaverledger/releaseControl/weaver` in the existing
Weaver project. It records the owner, full commit, source-manifest hash, starting
revision/image and operational aliases. Conditional Firestore writes allow one
claim to win across independent checkouts. This does not build or deploy, alter
IAM, or change the established Cloud Build / immutable-image release routes.
It uses the same existing Google/clasp connection as those operational tools;
it does not launch sign-in or request more scopes.

Before a build or Console trigger, from the clean approved source checkout:

```sh
node scripts/release-lock.mjs acquire --commit FULL_APPROVED_COMMIT --owner OWNER_EMAIL --receipt /private/tmp/weaver-release-claim.json
node scripts/release-lock.mjs check --receipt /private/tmp/weaver-release-claim.json
```

Keep the private receipt through the existing build, canary and promotion.
Run `check` immediately before each mutating phase. It refreshes main and rejects
changed source, base image/revision or operational aliases. Use `renew` before
the one-hour lease expires if the verified starting production state still
holds. The claim remains blocking after expiry: there is no automatic takeover.
For the existing shell source-deploy route, set `WEAVER_RELEASE_RECEIPT` to this
path. That script now requires a valid claim and builds with no main traffic;
it no longer promotes an unspecified `latest` revision.

After exact-revision promotion and verification, use:

```sh
node scripts/release-lock.mjs finish --receipt /private/tmp/weaver-release-claim.json --verification /private/tmp/weaver-release-verification.json
```

The verification receipt must bind `sourceCommit` and `sourceManifestSha256` to
the claim, `revision` and `imageDigest` to production, and contain actual positive
`ok`, `sourceImageVerified`, `authenticatedCanaryVerified`, and
`protectedRecordsVerified` checks. Main and both operational aliases must agree.
Closing also checks for active Cloud Builds in the existing regional/global
Weaver locations. A plan, build completion or healthy homepage is insufficient.

If the release was not promoted, `abort --receipt PATH` closes only the exact
owned claim after checking that its starting image/revision/aliases remain
unchanged and no build is active. A changed deployment or lost receipt needs
reconciliation, not deletion or a forced unlock. `status` is read-only. The
intended claim is saved locally before its conditional write, so a lost response
can be compared with the current durable claim without trying a second write.
`--source-root PATH` allows an operator to run the guard from separately installed
tooling while still checking the exact clean source being released.

### Enforcement boundary

The maintained shell entry point requires the shared claim. Console triggers
and immutable-image API operations must use the same acquire/check/finish
protocol. Direct cloud writes and old private one-off scripts can bypass this
coordination unless their operators adopt it. This is a tested shared gate, not
an IAM restriction or proof that every historical caller has been migrated.
Do not call universal enforcement complete until all active callers use it.

Book-intake acceptance: `node scripts/check-book-intake-release.mjs --season`
checks the configured season together, reports all missing books, and verifies
distinct keyed choices plus one readable keyed preview per available book.
It makes no submissions and never treats publishing metadata as readable text.

## Save and delivery recovery

A review save validates the entire batch's targets before writing. Stable source
IDs must match their current rows; the current UI also supplies the displayed
content and decision. Legacy row-only sources require that content snapshot.
Caption-only saves disable QI and INT on the server. After the sheet write, the
server reads back source identity and all requested values before updating the
ledger or queuing delivery. A failed ledger sync is reported separately from a
saved sheet decision and cannot queue downstream delivery.

Sheets values writes do not provide a compare-and-swap transaction. The checks
catch observed row movement and mismatched readback; they do not guarantee that
an independent sheet reorder cannot race the write. An uncertain result remains
unresolved. Do not claim that the wrong-row race is mathematically eliminated.

For an uncertain save, first read the exact source ID, current source row,
decision, graphic flags, corrections and source ledger. Check the delivery
ledger separately. Do not infer receipt from a disappearing queue row or repeat
a whole batch because a response was lost. Reconcile only the failed stage,
retaining canonical IDs and any existing attempt/receipt. A received or
acknowledged delivery is not eligible for an automatic resend.

In Needs correction, Return to review saves the explicit overrides and clears
correction/approval state. It does not approve the excerpt or queue graphics or
PP delivery. It is accepted only for a source currently in correction. Pending
and approved reads use the saved overrides while retaining original source
cells. Normal review's No decision continues to leave a record unchanged.

Sign-in becomes available after bootstrap loads its settings. Expired or failed
sign-in does not refresh queues, erase draft inputs or replay a mutation. Sign
in again, inspect the exact saved state if the earlier outcome was uncertain,
and submit only the changes that remain necessary.

## Recovery during a release

Keep the serving revision, immutable image, operational aliases and protected
record snapshot from immediately before the release. Before promotion, a
failed build or canary leaves main on that verified base. Remove only the
release's temporary job/tag when appropriate; preserve operational aliases.
Once no build remains active and the exact original base/aliases still serve,
`abort --receipt PATH` can close the owned claim. Never delete a stale claim to
make another release proceed.

If verification fails after promotion, keep the release claim held and inspect
the actual source/image/traffic and record differences. Do not use `finish`
with fabricated successful checks. The pre-promotion `check` intentionally
rejects changed main traffic; it is not an automatic rollback command. An
operator must reconcile the owned claim and current resource version before
restoring its captured, verified base revision and matching operational aliases
through the existing Cloud Run route. Verify the restored main/aliases and
protected records, confirm no build remains active, and then abort the claim
only if its exact original base has been restored. Otherwise leave it held for
explicit reconciliation. Never roll back approval or receiver data along with
application traffic.

After a claim has already finished, prepare a normal tested revert commit on
current main and release it through a new claim. Do not mislabel an older image
as having the current source manifest. Release rollback does not authorize
resending content, clearing holds or deleting receiver history.
