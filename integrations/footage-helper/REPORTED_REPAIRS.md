# Reported footage repairs: Eric Tu, 2026-11-04

Status: **staged; production execution blocked pending a specific exception to the user's new-sheets-only rule**. No listed repair has been executed by this work. Symone's earlier fix was independently verified as already complete. Automatic approval review rejected the repair execution because it targets the existing production inventory. Do not retry through another tool while that approval is pending.

Source: Sam's signed-in Gmail browser, thread “Footage Management Fixes,” five messages from Eric Tu (`eric61tu@buttonpoetry.com`), Oct 2, Oct 6, Oct 7, Oct 20 and Nov 3, 2026. The connected Gmail account was Emily's; its search was not used as evidence about Sam's inbox. Thread: https://mail.google.com/mail/u/0/#all/FMfcgzQhWfXrsZpvZRPkVDTggCrnSHpn

## Exact staged work

| Case | Current evidence | Planned action |
| --- | --- | --- |
| Donte September 2017 | Six Add Title rows 6647–6652 contain the Center Cam folder instead of a video URL. Exact clips 17–22 were found in that folder and verified by Helper for name, video MIME type and parent. | Set AW6647:AW6652 to those six resource-key-preserving file URLs; fill blank filename cells E4526:E4531 with the verified names. |
| Symone | CH4296 is blank; one correct Symone Johnson-Hawkins registry entry; all six title rows use the correct name. The Oct 6 Eric Repairs Log corroborates the earlier fix. | No change. Already corrected. |
| Garner | Legacy registry source C30 still contains Krystal Garner; canonical Garner exists at CH4298. Ten older title labels still use the historical name/content IDs. | Clear only C30 to remove the duplicate selectable name; retain historical title IDs and record the alias-migration follow-up. |
| Kyle Tran Myhre | Unique file ID `0B_-uqSs_rEYCMFlQT1FpVnJDcDA` identifies source row 6705, “The pen on the other side of the kitchen table will not move.” AT6705 says No video release. | Record Eric's explicit author/feature/workshop contract category in AT6705. This records the reported category; it does not independently certify publication clearance. |
| Mariam Tadesse | CH6822 and four source title rows use Miriam. Four derived rows have no FV lookup entries or scores. | Change CH6822 and L6823, L6837, L6839, L6842 to Mariam Tadesse; verify four derived names/content IDs and preserve media links. |
| Joel Bauman | CH6974 contains Joe; CH6975 and CH6976 contain Joel; the one title uses Joel. | Clear only CH6974. Keep the historical row, correct entries and title. |

Total: **20 literal cells — 14 source fields and six script-owned filename cells**. No formulas, spreadsheet rows, Drive files, event closure state or prior scores are changed. Refresh only the two canonical performer selectors (1479705159 and 1675652896) plus Performer List help text (1653265853), then verify parity with the formula-derived registry.

Prepared audit: https://docs.google.com/spreadsheets/d/19EGO_7NooedZ0uRkW7gic2KuwKrITGJSiMqIxdaEo4o/edit
Helper report: test Helper `?page=reported-repairs` (owner-only, read-only).

## Execution, evidence and recovery

`EricRepairs.gs` accepts no browser-supplied cell addresses or replacement values. It rebuilds this fixed plan from current records and Drive evidence, rejects formulas and unexpected before-values, fingerprints the preview, logs INTENT in a separate Helper-owned file, then applies only its listed cells. It verifies source and filename invariants excluding the listed cells, preserves formula anchors, checks derived names/links/rights and refreshes only the performer selectors/help. Unrelated event data is retained as hashes rather than copied into the audit.

Writes are sequential Apps Script writes, not an atomic database transaction. A failed write or verification is UNCERTAIN, blocks a repeated execution and requires readback. `recheckEricReportedRepairs` reads and verifies without repeating the source mutation. The owner-only undo restores the recorded before-values only if the relevant source invariant holds and each target still contains its recorded before/after value; changed records are preserved. The executor/undo paths have unit tests; production application and undo have not yet been exercised.

The preview was verified live on test Helper v55. The normal employee Helper remains v43. Source/form execution was rejected before its process ran. This is unrelated to authorization of the separate PAGES curation-link pilot, which remains disabled.

## Additional findings needing review

1. Donte: source row 6458 already references clip 17. The six planned folder-link corrections would leave a repeated clip-17 reference and no clip-15 reference. Do not change 6458 without checking the actual content/bout evidence. Link repairs do not establish event completeness.
2. Mariam: source 6823 has a folder URL and the title B****; source 6842 has an actual file with the same title. Check whether the first entry should be repaired or merged. No title/row deletion is authorized or implemented here.
3. Garner: introduce an explicit historical alias before migrating ten older content IDs; do not silently break FV/score joins by relabelling all historical title rows.

## Practical roadmap: move the work into Helper

1. **Report from the event/title.** Add a “Report a problem” action that pre-fills event, source row and Drive IDs. Eric/staff choose an issue type, describe the correction and attach evidence. Email can remain evidence, but the case/status/owner live in Helper. Start with wrong media link, performer spelling/duplicate and reported rights-category corrections.
2. **Catch the common errors at intake.** Check that a per-video link resolves to a video rather than a folder; populate/check the filename; flag repeated exact video IDs within the same event/version role; flag approved performer aliases and likely duplicate names. Use the existing maps and folder context rather than rescanning Drive per case. Similar names are suggestions, not automatic identity merges; OM/FV and horizontal/vertical versions stay distinct.
3. **Let the video team resolve ordinary cases.** Show a small preview of the old/new value and affected records, permit the narrowly supported fix, read back its effects, and retain the audit/undo history. Give Eric the appropriate reviewer permission so Sam is needed only for ambiguous identity, rights evidence or unsupported changes. Expand automation from repeated successful repairs rather than building a general spreadsheet editor.

Implemented now: fixed repair executor, separate audit, read-only Helper report and unit tests. Still roadmap work: general staff issue intake, assignment/review queue, reusable detectors, staff repair permissions and automatic repair policy. No notifications, recurring automation or email replies were created.

Deployment integration: preserve the current Helper project and add EricRepairs.gs plus the explicit owner-only POST cases in WeaverBridge.gs. In the existing doGet, before the footage route, add:

```js
if (page === 'reported-repairs') {
  weaverRequireOwner_();
  return HtmlService.createHtmlOutput(ericRepairPage_()).setTitle('Reported footage repairs');
}
```

The read-only report was verified in the browser with “Staged — no production changes applied,” all pending corrections, the separate audit and the follow-up findings. The audit remained empty after the rejected execution.

Evidence correction: the complete AU:AW scan of source rows 6437–6458 shows clips 1–14, then clip 16 at row 6457 and clip 17 at row 6458. Thus the gap after the staged 17–22 replacements is clip 15. Earlier provisional references to a clip-16 gap were incorrect and were corrected before execution.
