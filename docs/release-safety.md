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

## Still open: one shared release lock

The source guard is implemented; a global release lock is not. A local checkout
lock cannot protect independent chats/checkouts. Wire a single conditional
durable lease into the existing shared release runner, with a named owner,
approved SHA, captured serving revision and explicit release after verification.
Do not introduce a second deployment service or silently force-delete an
abandoned lease. All release entry points must honor it before claiming that
concurrent-release protection is live. This needs a bounded release-path change,
not a repository split.

Book-intake acceptance: `node scripts/check-book-intake-release.mjs --season`
checks the configured season together, reports all missing books, and verifies
distinct keyed choices plus one readable keyed preview per available book.
It makes no submissions and never treats publishing metadata as readable text.
