# Audited event curation-link registration

Updated 2026-11-04. Initial adapter: **Helper-owned sandbox only**. The normal employee deployment and production inventory are not enabled for this action.

## Intended workflow

After Helper confirms the current inventory manifest has a saved, complete Weaver receipt, preview the event and its specific task URL. Register once, then read back both the source record and the effective Curation view. Keep existing Google Form registrations intact. Opening or registering a task does not establish that curation has happened or that the event is complete.

The test Helper exposes preview, register, withdraw, recheck, and history. History survives a browser reload and links to the complete Audit worksheet. The first prototype is owner-only. The server reconstructs event/link data; the browser cannot choose a destination workbook or supply a replacement link.

## Paper trail and failures

Every operation has a UUID, actor, timestamp, source inventory/event rows, Weaver receipt/set identifiers, event/title snapshot, destination ID, and the prior registrations. INTENT is persisted before a source write. Later entries record source location, after values, verification outcome, or an error. The application appends audit entries; withdrawal never deletes them. This is an operational audit in a Google Sheet, not tamper-proof storage against someone with spreadsheet edit access.

A script lock serializes this Helper's operations. A fingerprint rejects a stale preview. A repeated click preserves an existing registration. Any incomplete operation blocks another append. Recheck reads the owned record and Curation projection before recording a resolved status; it does not retry the mutation. If no unique outcome verifies, it stays unresolved. Audit entries over 45,000 characters are refused, including before the first source write. The initial adapter is intended for small event pilots.

## Withdrawal

Withdraw only a verified Helper-owned registration with an unchanged, unique source record and no conflicting active registration. Preserve all source fields, changing only the test action from `Add Curation Form` to `Withdrawn Weaver curation link`. Confirm the event/link no longer appears in the test Curation projection. Log intent, before/after and the verified result. Do not erase source rows or audit history. Do not undo reviews, scores, Poetry Please transfers, or calendar activity. History and withdrawal remain available even if the current Weaver receipt later becomes stale.

## Why the production adapter is separate

Authoritative bound inventory Apps Script inspected read-only:
`1ospLmrS5B4elcCFsyZHn7a-i9GPhvlL3F4ytm6Ec7tsCNX7wpy98Cgv7`.
The live `onFormSubmit` branch for Add Curation Form also invokes performer updates, event-close pickers, FV lookup indexing and closure bookkeeping. Calling the entire branch is broader than a link registration.

Production source uses action B, email AV, event AX and URL AY in Form Responses 1. Curation derives its event/link pairs from Add Curation Form rows. The test reproduces this projection in its own spreadsheet; it does not establish all production downstream effects. `FormResponse.submit()` does not fire installable submit triggers, and deleting a Forms response does not remove its external sheet copy. Sources: https://developers.google.com/apps-script/guides/triggers/installable and https://developers.google.com/apps-script/reference/forms/form.

Before enabling production, isolate the required picker refresh and any other actual dependencies; define the treatment of retained Forms response data; prove both registration and withdrawal across the production-equivalent derived state. Do not modify production formulas, invoke the full submit handler, or expand to a general source writer. The current code contains no production adapter.

## Verification

Run `node --test tests/curation_registration.test.mjs`. It covers the lifecycle, duplicate clicks, failed intent, stale preview, existing production link protection, lost append response, failed audit after append, projection failure, changed rows/conflicts, uncertain withdrawal and destination mismatch.

Live proof must additionally record deployment version, owned test workbook URL, source and view row counts, audit states, duplicate-click outcome, and final withdrawal readback. Until recorded, deployment/live verification remain pending.

## Live sandbox proof — 2026-11-04

Test Helper version 50 completed the PAGES registration → duplicate click → history → withdrawal → repeated withdrawal cycle. Operation `ddd147d3-1143-4d5c-aeed-3130060fa059` has exactly six audit entries. Independent Sheets readback confirmed one retained source row, action `Withdrawn Weaver curation link`, original event/URL intact, and zero effective Curation rows. MPMU preview preserved its existing Google Form link without writing a registration. Normal employee deployment remained version 43.

[Owned test spreadsheet / audit](https://docs.google.com/spreadsheets/d/1oEl-qQ4LpKnwuoBamEk8Dzhcwpl_BdDmqQdxS_knm0g/edit). Runtime failures were exercised with injected stores in local tests; they were not deliberately injected into the live Google service. Production derived-state reversal is still unverified and production registration is disabled.

## Derived-state integration — 2026-11-04

Test Helper v52/v53 adds `CurationRegistrationIntegration.gs`. It is a synthetic-only harness with three fixed events, the relevant current inventory formulas, and a newly created Google Form whose responses are closed. It does not copy production event or rights records. Automatic approval review rejected a full production-workbook copy into the shared test folder; that copy was not made.

Verified live:

- Open synthetic event: registration removes the event from both needs-curation lists, adds its URL to Curation, adds the event to Event Worksheet J and the test close-event picker, and preserves the closed-event list. Withdrawal restores the exact prior snapshot.
- Closed synthetic event: registration adds its link, but never makes it available to close or reopens it. Withdrawal restores the prior snapshot.
- Existing-link event: preview preserves its link; no registration is written.
- Duplicate register clicks: no second source row or audit operation.
- Injected failure before the test form setter: the source remains saved and the operation becomes unresolved. Repeated registration and read-only reconciliation do not duplicate or finish the mutation. Explicit `repairDerived` records a recovery intent, refreshes the picker from the worksheet, and verifies the result without appending a source row. The same scenario was tested and recovered after withdrawal.

Final independent Sheets readback: three retained Helper-created source rows are withdrawn, one baseline legacy link remains effective, the closed-event list is unchanged, and 26 audit entries remain across the three operations (8, 8, and 10). This is proof against the relevant formulas and a real test Form, not proof of a production write or of every inventory consumer.

Fixture workbook: https://docs.google.com/spreadsheets/d/1HWfo6NGv7MIsSJOAusHocB7OH8C5JbBNuyKesTsI27U/edit
Fixture Form: https://docs.google.com/forms/d/1fm9iGoTU7kZnjsGPtgBaRT1y8ZUuBIMdiWO5dqN0baE/edit

## Current production preview and next slice

The owner can use **Preview live inventory change (read only)** in the test Helper. It reconstructs the current Weaver connection, checks existing source registrations, compares the two needs-curation lists, checks event closure consistency, and compares the canonical close-event picker with Event Worksheet J. It returns the planned source fields and effects but never appends a row or calls a form setter. Production writes remain disabled; the shared spreadsheet store still rejects the production ID.

Current live results on 2026-11-04:

- PAGES footage from Rudy Francisco 2026: current Weaver receipt; no existing curation link; open; present in both needs-curation lists; not ready to close; picker matches the worksheet. No preflight blockers. This is a candidate for the first narrow production pilot, not an assessment of event completeness or curation approval.
- Vertical Versions - MPMU 2026 Finals: already has a curation link and is closed. Preserve its registration and closed state; do not replace its existing link as part of the pilot.

Remaining work: implement the one-event production adapter with its audit in a Helper-owned file, then conduct an explicitly scoped PAGES pilot. Append only the timestamp/action/email/event/URL source record; refresh only close-event picker 466810607 from the existing formula result. Do not modify formulas, execute the broad submit handler, close an event, or alter existing registrations. These Helper-origin records would be source-sheet actions with their own audit, not Google Forms responses; document that distinction rather than claiming Forms response deletion reverses them. Production withdrawal must preserve history and verify every affected view/picker as the synthetic harness now does.

## PAGES-only production adapter prepared — test Helper v54

`CurationRegistrationPilot.gs` now implements the narrow adapter. It remains **disabled and unapplied**. The backend permits only the exact PAGES event and the deploying owner. It requires authorization of the exact fresh preview before registering. A second append is prohibited once the pilot has an INTENT, including after withdrawal. Current receipt validation is required to register, while history and owned-record withdrawal do not depend on a later receipt remaining current.

The source action appends only timestamp A, action B (`Add Curation Form`), actor AV, exact event AX, and task URL AY; all other source fields are blank. The only form mutation is refresh of close-event item 466810607 from existing Event Worksheet J. It does not create a Forms response, invoke the general form-submit handler, edit formulas, update FV indexing, close an event, change a score, or send anything to Poetry Please/calendars.

The separate Helper-owned audit stores the intent, authorization, source before/after, exact operation ID, derived readback and outcomes. It stores hashes of unrelated event lists instead of copying those lists into the audit. Withdrawal requires the unchanged Helper-created source row; it changes that row's action, preserves the original fields/history, refreshes the same picker, and verifies restoration. Unrelated inventory/picker changes cause refusal or an unresolved outcome; do not claim automatic reversal is guaranteed after other work has changed those dependencies.

Verified live, without authorizing the pilot: preparation returned `productionEnabled=false`, `canRegister=false`, and zero history entries. A register request was refused before a production write. Independent readback confirmed zero audit records and no PAGES curation-source row in AX:AY through the current source bounds (7,002 rows). 45 focused tests pass. Production registration/withdrawal themselves remain unverified until the scoped pilot is approved and executed.

Prepared audit: https://docs.google.com/spreadsheets/d/1bzvyiq8p2sGqJiWTTzSWo9lOvmIsVjyrug9VvNW_dwo/edit

Execution after approval: prepare a fresh preview; authorize its hash; register once; read back source, audit, Curation, both needs-curation lists, closed state and close picker. A lost response requires history/readback before retry. The owner can inspect the separate audit through **Show live registration history** in the test Helper. Withdrawal and derived recovery are owner-only backend actions in this MVP; normal employee Helper remains v43.

Approval needed: the user's earlier instruction restricts writes to newly created sheets. The first live pilot therefore needs a specific exception to append this one registration in the existing inventory and refresh its live close-event picker. Do not infer approval from the existence of the adapter or its deployment.
