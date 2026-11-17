# Verified platform output selection

The video pipeline first requests its existing 1080p profile for the current
finished source. A missing-profile 404 then opens the exact canonical parent's
rendition list. It can request a verified 1080p or 720p profile from a source
version already explicitly linked to the same performance. It never associates a
new recording by filename similarity or chooses an unverified master reference.

Profile preference is 1080p then 720p. Within a profile the current representative
is preferred. Multiple equally preferred verified assets require a choice; no
filename/id sort invents one. Source/canonical/profile/version, active state,
verified byte count and SHA-256 must agree with the listed selected asset.
Authorization errors, other receiver failures and crossed identities stay errors.
Missing or ambiguous verified outputs remain media-pending.

This is media evidence only. Existing editorial/release restrictions remain in
force. Calendar readiness requires the receiver's explicit, fresh, output-bound
approved-caption contract; neither a stored source approval nor a transcription
alone sets it. Transcription can corroborate words in a checksum-bound output
while reusing an existing human-approved excerpt; it does not create an approval.

The Calendar consumer must reconcile existing canonical/platform placement and
Post Record IDs before staging. Already scheduled/published history does not
prove the new certificate-based path, and the new readiness feed must not create
another placement for the same work just because its output resolution changed.

Regression coverage: `tests/video_rendition_selection.test.mjs`, existing curation,
pipeline and runtime-module coverage. Private per-output/cohort evidence belongs
in the Codex shared drive, never this repository.
