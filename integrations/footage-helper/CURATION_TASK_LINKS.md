# Event-specific curation links

The curation task identity is the existing Weaver `prioritySetId`. Share `https://<verified-weaver-origin>/?videoSet=<prioritySetId>`. It opens Excerpt gathering → Video and loads that exact event, with the current reviewer's progress. Opening a link does not write a review, decision, inventory submission, or publication.

The test Helper returns `curationTaskUrl` after a saved connection passes its current checks. Use **Check Weaver receipt**, then **Open this event’s curation task in Weaver**. Preview-only, incomplete, and inventory-changed results do not advertise a ready task link. The normal employee Helper deployment is separate from this test deployment.

Examples on the reviewed test bridge:

- MPMU: `https://footage-bridge---weaver-iasomqy6ja-uc.a.run.app/?videoSet=mpmu-2026-finals`
- PAGES: `https://footage-bridge---weaver-iasomqy6ja-uc.a.run.app/?videoSet=pages-matam-san-diego-2026-labc`

Unknown or malformed event IDs show an error; the tool does not fall back to a different queue. Completed reviewers see the completion state. Links carry no reviewer email or credential. Existing authentication and review-save rules apply. The playlist input also accepts same-site Weaver task links and retains Google Form support. Links to a different Weaver origin should be opened directly there, preserving deployment boundaries.

Inventory's `Curation!A2` and `B2` formulas read Add Curation Form submissions from response columns AX and AY. The URL column is passed through without parsing a Google Form ID. A Weaver task URL can therefore be represented in this read model using the existing event/link shape. This change does not submit or edit any production inventory record, alter readiness formulas, or change the live form's validation. Browser acceptance of the URL by that form remains a separate verification; do not claim that a production Add Curation Form submission was made.

An event-level link creates a destination for curation; it does not assign a named reviewer, certify event completeness, or approve publication. Preserve existing curation/review history when using it for already-reviewed events such as MPMU.
