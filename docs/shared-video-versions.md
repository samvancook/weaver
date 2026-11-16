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
