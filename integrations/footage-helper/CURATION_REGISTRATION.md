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
