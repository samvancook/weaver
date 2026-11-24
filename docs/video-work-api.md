# Current event work and direct assignments

`GET https://weaver.buttonpoetry.com/api/intake/video-work` is the read-only
machine contract for discovering current video curation and eligible excerpt
approval work. It uses the existing hosted library, saved source excerpts and
release-eligible pending review queue. No browser session, chat, workspace file
or local script is needed to read it. It introduces no write grant.

```sh
curl --fail --silent --show-error https://weaver.buttonpoetry.com/api/intake/video-work
```

Check `ok`, `schemaVersion`, `checkedAt` and `revision`. Responses are no-store;
the existing source cache is at most five seconds. Failed source reads return
503 rather than a fabricated zero-work report. `unassignedPendingExcerptCount`
reports ambiguous or unmapped sources instead of assigning them by poem title.

`events` gives each stable event ID, included performance-group count, collected
excerpt count, remaining reviewer slots, missing-excerpt coverage, pending
candidate counts, existing approved-excerpt coverage and pipeline stage counts.
Formats/resolutions explicitly belonging to one performance count once. Source
excerpt IDs count as collected excerpts; exact pending source rows count as
pending candidates. These are distinct measures, not a promise of unique text.

`tasks` contains required curation and eligible excerpt approval assignments,
sorted by affected video count and stable task ID. `largestTask` is the first
entry or null. `workUrls` point straight to the hosted task. Optional extra
candidates on a video that already has an approved excerpt are reported as
`additionalPendingExcerptCount`; they do not become required approval tasks.
When an approval assignment needs over 50 exact file IDs, links are split to
respect the existing assignment limit.

The event curation URL uses `?videoSet=EVENT_ID`. A fresh visitor can read the
linked event without a preloaded event picker or personal session. Entering a
work email loads the personal review queue. Task links retain their event; an
unknown task does not substitute All events. Saving still uses existing Weaver
controls and permissions. These links do not approve, send, edit, or publish.

Curation completion is separate from score eligibility, release restrictions,
verified media, receipt reconciliation and Calendar readiness. Those blockers
remain visible in event counts; `tasks` is not an inventory of all downstream
engineering or publication work. Receiver implementations stay in their own
projects and use their existing handoffs.

Acceptance requires a live unauthenticated read, an independent-machine read,
and browser verification that the returned event URL opens that exact work.
Local tests and this document alone do not establish deployment or acceptance.
