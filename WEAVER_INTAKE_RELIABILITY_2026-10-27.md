# Weaver intake reliability — 2026-10-27

## Verified current state

- The live Weaver video dashboard has six sets and 119 videos. The five priority sets remain under human selection; the top BPL card still shows `Choose`.
- The source Excerpt Tool sheet contains the WG delivery row and newer book excerpts. A transient apparent gap in a parallel range scan disappeared on a same-row read, so it is not evidence of a partial row.
- A 2026-10-26 intake request returned HTTP 500. Available Cloud Run request logging does not show its application exception or prove why a prior row had missing front columns.
- The Cloud Run service uses `weaver-runtime@button-weaver-internal.iam.gserviceaccount.com`, which has `roles/datastore.user` on its project.

## Changes in this branch

- A Firestore claim keyed by excerpt record ID protects book submissions from simultaneous app instances. A failed or unresolved claim prevents an uncertain Sheet append; an interrupted claim can be recovered after ten minutes with a version precondition.
- Book append results and repairs of partial front columns emit structured logs without excerpt text. The Sheet append has a 30-second timeout, and the existing post-write readback remains in place.
- `/api/intake/options` reports whether each publishing-order book is catalog backed, metadata only, or unverified during a Catalog fallback. The release checker reports Spring 2027 statuses.

## Open gates

- Confirm the Firestore claim behavior in the deployed runtime with a real approved excerpt before calling concurrent-write protection production verified. Do not use a synthetic quote.
- The original cause of the partial front row remains unknown until a future event yields the new append and repair logs.
- Catalog owns the supported book-list/status endpoint. Weaver still uses the current `%` search with a limit of 1000 until that API exists; replace the adapter only against a confirmed Catalog contract.
- A human must choose priority videos within each event. Transfer only chosen videos with a final playable asset and release clearance; then verify the Poetry Please record and calendar item by canonical ID.
