# Inventory receiver curation rights

The receiver uses the same curation eligibility as Footage Helper:
Have video release, the existing author/feature/workshop contract label and
Title not included on Video Release are included. No video release and
Performance Opt Out are excluded and explicitly reported by source-row identity.
Unknown labels remain blocked for review. Case/whitespace normalization is shared
between intake and exact-file event matching.

Excluded rows do not require file access or matching into an existing queue.
Their rights and original metadata remain in the manifest and excluded receipt;
only accepted assets enter a queue. A wholly excluded batch never reports a
complete transfer. Matching never uses excluded files to select another event.

This changes curation eligibility only. Original rights values, event/digest
identity, existing stored receipts/reviews, configured event rules, publication
restrictions, media roles and closure behavior remain unchanged. Preview stays
read-only and does not persist an intake. A later explicit import is separate.

Verification: mixed and wholly excluded batches, unknown rights, unreadable and
invalid included assets, existing queue reuse, publication separation and retained
legacy behavior are covered by the focused receiver and integration tests.
