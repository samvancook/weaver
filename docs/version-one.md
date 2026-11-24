# Weaver 1.0 stabilization

Scope frozen by the owner on November 24, 2026. Finish the existing product;
video-file mapping is the only parallel workstream. Keep one Weaver application
and repository. PP, Calendar, Catalog and PIG own their implementations.

## Acceptance list

An item closes only with current evidence for the exact records and operation.
A passing test, a healthy endpoint or absence from a queue is not alone proof of
successful delivery. Keep private record evidence outside the repository.

- [ ] Exact-record saves: approval, correction and Graphics QC target the selected
  source; retries preserve identity and do not create duplicate downstream work.
  Identical excerpt text must not implicitly approve or reject another record.
  Caption-only approval must not queue either graphic lane.
- [ ] Book intake: existing Catalog-backed books offer distinct keyed poem choices
  and readable previews. Metadata-only books report that limitation accurately.
  Exercise an authorized intake through durable readback; preserve existing keys.
- [ ] Excerpt/graphics pipeline: verify an authorized save, durable source decision,
  delivery receipt and correct QC identity. Reuse existing approvals and assets.
- [ ] Video pipeline: retain 8.5/five independent reviewers; reuse finished edits
  before requesting editing. Only verified OM with a complete negative finished
  lookup may create an editing request. Verify exact video/rendition/caption
  receipts and downstream readiness without changing publication history.
- [ ] Staff recovery: assignment links preserve scope; sign-in expiry or a failed
  save retains the user's work and gives an actionable retry/reconciliation path.
- [ ] Release/recovery: test suites pass; use the existing shared release claim,
  exact source, no-traffic canary and protected-record readback described in
  [release safety](release-safety.md). Record rollback and receiver reconciliation
  instructions against the actual release.

## Boundaries

No new dashboards, permission redesigns, services, repository splits, content
types or broad historical backfills for 1.0. Content exceptions need a specific
owner and next action; they do not automatically require a new feature. Closed
content stays closed. Coordinate receiver work through existing shared-drive
handoffs. Do not create synthetic editorial approvals to pass acceptance.

## First repair

Expose each excerpt occurrence independently, retain duplicate warnings, remove
implicit peer rejection and verify caption-only graphic flags. Local regression
coverage is implementation evidence; deployed save and receiver readback remain
separate acceptance steps. Update the shared current-state register with actual
release and acceptance receipts rather than marking this checklist from a plan.
