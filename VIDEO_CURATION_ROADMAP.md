# Weaver Video Curation Roadmap

## Immediate: Priority Video-Set Pilot

- Make the four approved 2026 Drive sets reviewable in Weaver: BPL Charm City, MN Writers Respond @ The Loft, MPMU Finals, and Ollie Schminkey Action Cam.
- Treat the Drive file ID as the stable source-video identity and retain exact folder/file provenance.
- Persist reviewer ratings independently from excerpt rows so a reviewer can rate and advance without selecting an excerpt.
- Use `dislike`, `meh`, `like`, and `moved_me` as the primary rating contract.
- Keep raw footage links and reviewer identities in Weaver. Send only approved excerpts, publishable media, and appropriate curation fields to Poetry Please.
- Enrich from Footage Inventory when a match is available without blocking initial review.

## Later: Historical Curation and Video Backfill

- Normalize the 2018-2026 quarterly curation responses into reviewer-level durable review records.
- Reconcile those records with the historical Curation Database `All_Responses` archive and `Scoreboard` aggregates.
- Join exact video assets through Footage Inventory identity and use the existing Looker export as a delivery and verification artifact, not source truth.
- Preserve historical 0.0-10.0 values and derive categorical compatibility without overwriting the original score.
- Backfill approved video/poem curation aggregates into Poetry Please without sending unedited footage links or reviewer PII.
- Produce matched, ambiguous, duplicate/version, missing, and unresolved reconciliation counts before retiring old operational forms.
