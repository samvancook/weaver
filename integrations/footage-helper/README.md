# Footage Inventory → Weaver pilot

This change adds an event handoff to the existing Helper Intake QC page and Weaver Video curation. It is implemented and tested locally, **not deployed or verified live**.

The source Helper project is **Retitle Portal**:
https://script.google.com/home/projects/1aGmHhfSQujviWDRN7WY5mMJEVbG3t81O0MucZ0JWLN266vbvoNoWa0k1/edit

`IntakeQC.html` is the October 29 editor snapshot with one added Weaver section. `WeaverBridge.gs` is new. Do not push the entire Weaver repository into the Helper project. Re-read current Helper source and apply only this section and new file, preserving any intervening changes. The captured source was not proven byte-identical to deployed version 43.

## Behavior

1. A staff member opens an event in Helper's existing Intake QC report and chooses **Send event to Weaver**.
2. Helper reconstructs the current event and title records with its existing `getIntakeReport`; the browser does not supply the title payload.
3. Weaver verifies the signed Helper identity token, validates the manifest and checks each individual video with its own Drive identity. For this first pilot only titles marked `Have video release` are accepted; other statuses are listed as blocked, not silently dropped. This is not a new ruling on author contracts or rights.
4. Weaver saves a receipt in its `footageEventIntakes` Firestore collection, reads it back, and makes the accepted files available through its existing Video curation queue. Inventory performer and title fields take precedence over filenames. No camera-folder expansion or performance reconstruction happens.
5. Helper shows received/blocked counts and **Ready for closure review** only when every submitted title was accepted and current Drive metadata checks succeed. **Check Weaver receipt** rechecks Weaver access and compares current inventory title data with the accepted manifest.

A receipt confirms durable queue registration and metadata access. It does not confirm video playback, completeness of the original inventory work, or permission to publish. Every imported source remains publication-restricted. Media type stays unknown because an OM event folder alone does not prove that a linked title video is OM. Existing review and excerpt tools remain the curation surface. Event closure stays manual in the existing inventory workflow.

Retries retain one stable event ID and update access results for the same manifest. A changed manifest returns 409 and preserves the earlier queue; revision/update support is deferred until after the pilot. Partial sends appear with explicit blocked rows and cannot report closure readiness. No receipt or bridge code writes to production Footage Inventory sheets, renames/moves files, or deletes media. Helper stores only event receipt pointers in Script Properties.

## Configuration and rollout still required

Both sides are disabled until configured. No credentials or OAuth audience were invented or copied into this change.

- Weaver: `WEAVER_FOOTAGE_BRIDGE_ENABLED=true`, `WEAVER_HELPER_TOKEN_AUDIENCE` set to the Helper's exact verified OAuth client audience, and `WEAVER_HELPER_ALLOWED_EMAILS` set to the actual deploying owner's verified email. Existing Firestore production ledger configuration and Drive identity must be retained. The new collection uses the same Firestore client as other Weaver receipts.
- Helper Script Properties: `WEAVER_BRIDGE_URL` set to the reviewed Weaver HTTPS origin, `WEAVER_BRIDGE_ENABLED=true`.
- Helper authorization: inspect the project's currently resolved scopes and preserve them. Merge the additions in `oauth-scope-additions.json` into the manifest's `oauthScopes`; that file is an additions list, not a complete manifest. Authorize as the deploying owner. `weaverConnectionInfo_()` returns only the identity token's audience and email to assist configuration; it does not expose the token. Weaver independently verifies the signature/audience/email on requests.
- Deploy a preview Weaver revision and a test Helper deployment before updating the employee-facing deployment. Neither is deployed by this PR. Do not use the repository's default clasp target for the Helper.

Google documents the explicit `openid` scope and optional email claim for `getIdentityToken`: https://developers.google.com/apps-script/reference/script/script-app#getIdentityToken(). Scope configuration: https://developers.google.com/apps-script/concepts/scopes.

## Pilot and completion evidence

Current inventory reads on October 29 found **Adam Falkner NYC Release**, two Add Title records (rows 6138–6139), both marked `Have video release`, individual Drive MP4 links, and no matching closure entry in source Close Event responses or the closure worksheet. Two whitespace-equivalent Add Event rows (6107, 6137) point to the same OM folder. They map to one stable bridge identity. The second inventory title uses “Honest” while its Drive filename uses “Kindest”; this remains an observed metadata discrepancy, not a confirmed bad link. The bridge preserves the inventory wording and does not repair it.

The connector returned metadata for both MP4s. That is **not** evidence of Weaver runtime access, playback, or a successful import. The browser automation surface was unavailable during this implementation; live Helper authorization and pilot round trip remain unverified.

Before calling the pilot complete: submit the current two-title manifest through the test Helper, read the persisted receipt back, load the same event's two videos through Weaver's existing Video interface, verify playback as a curator, and recheck Helper's receipt and the unchanged open inventory state. Repeat the send and verify one event with two items. Do not close the event automatically.

## Validation performed

- 35 JavaScript tests passed across bridge contract, mocked route/queue rehydration, Helper server calls, authentication, and existing video handoff/release flows.
- 5 new Python ledger tests passed (receipt readback, retries, changed-manifest protection, atomic creation, no SQLite fallback); 6 existing curation claim tests passed.
- Server/client syntax checks and Python compilation passed.
- The wider dashboard test file has one existing failure: it expects `selectedExcerptRecordIds` inside a gate literal that current base commit `3d8ae6d` does not contain. The unmodified base has the same missing pattern; this PR does not alter that handoff literal or that test. Six other tests in that file passed.

The integration tests use mocked identity/Drive/storage services. They exercise production receiver, queue loading, and Helper functions, but do not prove production permissions, real Firestore writes, owner consent, or a live deployed round trip.

## Live-test connection setup

The rollout adds `ConnectionSetup.gs` and the small `Code-doGet.patch` to the current Helper source. The Intake QC page gains **Connect Weaver (owner)**. This accepts no arbitrary endpoint: it configures the reviewed `footage-bridge` preview and returns non-secret token audience/email metadata. Only the deploying owner can connect or read these details. The owner-only `?page=weaver-connection` route returns the same metadata so deployment verification can read it without copying a token or inventory data.

`appsscript.live-test.json` makes the existing Sheets/Forms access and the read-only Drive access explicit, and adds the OpenID/email/external-request scopes needed for this handoff. It preserves DOMAIN access and USER_DEPLOYING execution. No repair, mutation, trigger, or inventory function is run by connection setup. Apply it only to the separate test deployment first; keep the employee-facing version 43 deployment unchanged until the pilot succeeds.
