# Weaver Rules

## Handoff Authority

- Approved excerpt truth is the Weaver approved-excerpt export feed.
- Graphics lifecycle truth is the graphics handoff ledger.
- `data/weaver_runtime.db` is cache/scratch only, not durable authority.

## Required Publishing Metadata

- No excerpt may be passed downstream without both:
  - `releaseCatalog`
  - `bookShortener`
- Missing publishing metadata must block handoff and surface as a failure, not a silent send.

## approvedForGraphics

- `approvedForGraphics` means editorial eligibility/requested-for-graphics only.
- It must not mean:
  - graphic completed
  - QC passed
  - asset exists
  - P.I.G. finished

## Release-Catalog Lanes

- The queue lane previously framed as a seasonal/year lane should derive from publishing-order release metadata.
- A book belongs in that lane only when it has both:
  - `releaseCatalog`
  - `bookShortener`
