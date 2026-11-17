# Shared video versions

Format variants use one performance curation assignment with all original assets,
reviews and excerpt references retained. The initial explicit relationships in
`video_version_families.mjs` cover BPL Charm City / Click and Waterfall, including
historical reviewed-source IDs and current exports. Grouping never uses fuzzy
filename similarity to join unrelated performances.

The implementation supports any aspect ratio or resolution. New confirmed
relationships can use the same family contract; this release does not claim an
automatic census or reconciliation of every duplicate in Drive. Different edits
remain linked to the poem; their timing stays source-specific.

Queue, progress, reviewer export and candidate reads use the grouped file model.
Existing reviews are not rewritten. Their original source IDs remain in review
evidence; distinct reviewers are counted once and excerpt references are unioned.
Existing latest-review selection behavior is retained. The UI offers all linked
version links, and old version deep links resolve to the shared assignment.

Poetry Please receives optional `performanceContext` on its existing canonical
video through the additive audited context endpoint. Handoff verification checks
stored performance, assets and excerpts and confirms primary media is unchanged.
No new parent video, media upload, event closure or publication is implied.
Protected rendition APIs expose the same context to downstream consumers.

Regression tests are in `tests/video_version_families.test.mjs` plus the existing
queue, bridge, curation handoff and excerpt source suites. Real-data preview and
protected receiver readbacks are retained locally under
`.evidence/shared-variants-20261116` (not committed: contains private review data).

## Production verification — 2026-11-16

Weaver PR #98 and Poetry Please PR #400 are merged. Current verified Weaver
revision: `weaver-shared-variants-v2-20261116`, preserving the concurrently deployed
excerpt receipt-pacing fix. Canonical Weaver, `footage-bridge` and `video-queues`
all return 54 BPL performances. Click has three current assets, 12 independent
reviewers and 14 excerpt references; Waterfall has two, nine and 13 respectively.
Both existing Poetry Please parents retain the same primary media and expose the
same context through protected receipts and the downstream rendition API.

All 986 original review records, 44 curation gates and seven intake records were
compared to the complete before snapshot using typed-map equality: no changes or
deletions. The two context updates have receiver audit IDs
`G0rj3nlyduX3rnUsFp5j` and `IVX4iwkpc2PjHi2PE8gz`.

Regression results: 224 of 226 Weaver Node checks pass. The two failing multicard
project tests also fail on the unmodified base; they concern graphics project
identity/QC and were left unchanged. The affected video/footage/excerpt checks
pass. Poetry Please's production CI, authorization emulator, public health and
anonymous bootstrap checks pass. No publication or Calendar readiness was set.

## Graphics regression follow-up — 2026-11-17

The two previously failing graphics checks were stale tests. The legacy project
fixture now exercises the intentional request-identity behavior added in
`e73260b`: saved text edits are accepted for the exact request, while mismatched
content/request identities and uncorroborated text conflicts are still rejected.
The QC fixture now models a saved, verified editable-project identity. A new
negative test verifies that failed project revalidation withholds all editable
links while preserving the approved card set. No production code or data changed.
The complete Weaver Node suite now passes 227/227 checks, including video versions.

## Loft cross-folder format pairs — 2026-11-17

The Loft set's horizontal FV folder is separate from its vertical FV folder.
Twelve exact author/title pairs have equal current Drive durations, with 1920×1080
horizontal and vertical counterparts. Explicit families retain the existing
vertical review-source IDs. The loader now reads additional configured folders
once, then presents 24 files as 12 performances with shared review/excerpt context.
Inventory-backed sets retain their existing file list and do not scan extra folders.

Real-data preview preserves all twelve progress summaries exactly: seven unique
reviewers per performance. Eleven remain below the existing delivery threshold;
no score, release decision or priority was changed to force a send. Only Sarah
Ghazal Ali / Under Such Conditions currently has a Poetry Please parent. Its two
version links and six excerpt references use the existing additive context path.
The receiver's media is verified, reviewOnly is true, and its receipt currently
omits the caption/calendar readiness contract. This is not publication confirmation.

All 229 Node checks pass. Release source preserves the concurrent live
receipt-root repair, which was not yet on main when this change was prepared.
Private snapshots, exact file metadata, before receipts and link preview:
`.evidence/loft-20261117/`. No inventory cells, media or event closures are changed.
