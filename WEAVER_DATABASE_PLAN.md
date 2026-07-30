# Weaver Database Plan

## Goal

Replace the spreadsheet-backed graphics workflow with a Weaver-owned runtime database, one seam at a time.

This is the first pass of that plan:

- define the core tables now
- keep the live app stable
- migrate in phases instead of cutting everything over at once

## Recommended Production Shape

- **Production:** Cloud SQL Postgres
- **Local/dev modeling:** SQLite

Reason:

- the data is relational
- we need durable server-owned writes
- we need append-only completion and QC history
- we want P.I.G. and Poetry Please to read/write against stable ids

SQLite is useful here to lock the schema and build migration logic locally, but it should not be treated as the final production persistence layer for Cloud Run.

## First Database Slice

This first Weaver-owned data model covers the graphics pipeline:

1. `graphics_requests`
2. `graphics_request_items`
3. `graphics_completions`
4. `graphics_qc_reviews`
5. `poetry_please_handoffs`

This is the right first slice because it captures:

- what needs to be made
- what P.I.G. returned
- what QC decided
- what was handed downstream

without forcing us to migrate excerpt gathering first.

## Table Intent

### `graphics_requests`

One logical graphics request owned by Weaver.

Use this as the canonical record for:

- open request
- completed-returned request
- QC-approved request
- rejected request

### `graphics_request_items`

The child rows/excerpts that belong to a request.

This lets us preserve grouped requests without pretending every request is always a single excerpt.

### `graphics_completions`

One returned asset from P.I.G.

Important properties:

- append-friendly
- source-tool-aware
- asset-link-aware

### `graphics_qc_reviews`

Append-only QC decisions.

This is intentionally not “just a status column.”
We want history here so we can see:

- what was approved
- what was corrected and recreated
- what was rejected
- when and by whom

### `poetry_please_handoffs`

The downstream handoff ledger.

This is how Weaver can tell the truth about:

- queued for Poetry Please
- previewed only
- imported
- import failed

## Migration Order

### Phase 1: Schema + local runtime model

Done in this pass:

- schema file
- Python runtime helpers
- DB init script

Files:

- [db/weaver_runtime_schema.sql](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/db/weaver_runtime_schema.sql)
- [weaver_runtime_db.py](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/weaver_runtime_db.py)
- [init_weaver_runtime_db.py](/Users/buttonpublishingone/Desktop/CODEX/Excerpt%20Management/init_weaver_runtime_db.py)

### Phase 2: Dual-write P.I.G. completions

Next step:

- keep writing P.I.G. completions to `PIG - Completed Graphics`
- also write them into `graphics_completions`

This is the safest first dual-write because that flow is already server-owned and stable.

### Phase 3: Dual-write QC reviews

After completions:

- keep saving QC to the current sheet lane
- also append into `graphics_qc_reviews`

At that point, the spreadsheet stops being the only owner of QC truth.

Important rollout note:

- the first dual-write pass should cover **P.I.G.-backed QC rows first**
- legacy cleanup-sheet-only rows can stay sheet-native until we decide whether to model them as synthetic completions or retire that lane

### Phase 4: Read QC from the DB

Switch `Graphics QC` reads to:

- requests + completions + latest QC reviews from the database

Keep the spreadsheet as fallback only during this phase.

### Phase 5: Write graphics requests into the DB

When excerpt approval creates “needs graphics” work:

- create/update `graphics_requests`
- create/update `graphics_request_items`

That becomes the true source for P.I.G.

### Phase 6: Move P.I.G. reads off the sheet entirely

P.I.G. should read:

- open requests
- request items
- existing completion status

from Weaver-owned database tables, not from the helper sheet tabs.

### Phase 7: Poetry Please handoff from the DB

After QC approval:

- handoff records should be created from `v_qc_approved_graphics`
- Poetry Please should consume that dedicated feed
- handoff status should be recorded in `poetry_please_handoffs`

### Phase 8: Decommission spreadsheet graphics lanes

Once reads and writes are stable:

- `Queue - Needs Graphics` becomes optional transition tooling
- `Cleanup - Created Graphics` stops being the QC source of truth
- `PIG - Completed Graphics` becomes a temporary audit lane only, then removable

## Practical Next Steps

The next three concrete Weaver-side implementation tasks should be:

1. Dual-write `POST /api/pig/completed-graphics` into `graphics_completions`
2. Dual-write `POST /api/save-graphics-qc` into `graphics_qc_reviews` for `PIG - Completed Graphics` rows first
3. Add a database-backed `GET /api/poetry-please/qc-approved-graphics` implementation

That gives us the first real payoff:

- QC no longer depends solely on sheet columns
- Poetry Please can read from a Weaver-owned model
- P.I.G. completions become durable app state instead of a sheet-only lane

## Administrator Authentication Roadmap

Server-side authorization is required for staff actions that change Weaver or downstream
workflow state. Hiding controls in Under the Hood and limiting batch size reduce blast
radius, but they are not authorization boundaries.

### Bare-bones implementation

- Require a Google OAuth bearer token on administrative mutation routes.
- Verify the token with Google on the server.
- Require the token audience to match `WEAVER_GOOGLE_OAUTH_CLIENT_ID`.
- Require a verified email allowed by `WEAVER_ADMIN_EMAILS` or
  `WEAVER_ADMIN_EMAIL_DOMAINS`.
- Default the allowed domain to `buttonpoetry.com`; production may replace or narrow it.
- Cache successful verification briefly in memory and never log the token.
- Log the verified administrator email, HTTP method, and route.
- Keep contributor intake and read-only endpoints outside the administrator guard.

Covered staff mutations:

- excerpt handoff retry and approved backfill
- review saves and single-review saves
- graphics QC saves, link saves, rework creation, handoff retries, and folder-import apply
- stalled P.I.G. recovery changes
- Poetry Please repair sync and repair-status retry

### Follow-up hardening

1. Replace domain-wide access with a small explicit `WEAVER_ADMIN_EMAILS` allowlist or a
   managed administrator group.
2. Store the verified administrator identity in each durable mutation history record.
3. Add rate limits and structured security audit logs.
4. Add automated route-policy tests so new administrative mutation endpoints cannot ship
   without an explicit authorization classification.
5. Add separate service-to-service authentication for P.I.G. lifecycle endpoints. These
   routes must use workload credentials and must not depend on an interactive staff token.
6. Add a Poetry Please acceptance callback or readable returned/resolved status before
   Weaver automatically marks repair requests resolved.
