# Eric’s footage corrections — executed 2026-11-04

**Executed and independently verified.** Operation `31b3bf52-ae4e-4fba-ace7-1c4503ce7698` completed with VERIFIED, zero errors, and 20 changed input cells, all in **Form Responses 1**. Six generated filename outputs were populated by reading the new source-column O values and validating their column AW links. No formulas, rows, media, closure status or existing scores were edited.

[Production repair audit](https://docs.google.com/spreadsheets/d/19EGO_7NooedZ0uRkW7gic2KuwKrITGJSiMqIxdaEo4o/edit). The test Helper has an owner-only, read-only `?page=reported-repairs` report. Normal employee Helper remains v43; test Helper v59 contains the repair and rollback implementation.

## Authorization and source evidence

Eric Tu’s “Footage Management Fixes” thread in Sam’s signed-in Gmail contains five messages dated Oct 2, Oct 6, Oct 7, Oct 20 and Nov 3. Sender `eric61tu@buttonpoetry.com`, recipient `sam@buttonpoetry.com`, verified in the browser. [Email evidence](https://mail.google.com/mail/u/0/#all/FMfcgzQhWfXrsZpvZRPkVDTggCrnSHpn).

The initial 14-source/6-destination-cell plan was blocked by automatic approval review under the standing new-sheets-only rule and was not executed. The user then required undo proof, source-tab inputs, no further review presentation, and explicitly instructed “confirm and then execute,” specifying that filenames must be changed in Form Responses 1 and populate to the destination. The plan was revised to 20 source inputs plus a source-driven filename refresh; the isolated full rollback was proved before execution. This authorization did not enable the separate PAGES curation-link pilot.

## Every input column and destination checked

The live headers, current values, absence of target formulas and source-row actions were checked before execution. The executor permits only these exact 20 source cells; it accepts no caller-supplied replacement values or addresses.

| Source column | Verified live meaning | Exact cells changed |
| --- | --- | --- |
| C | Select Performer, legacy Add a Performer entry | C30 cleared: unused Krystal Garner duplicate |
| CH | If not listed, enter performer here | CH6974 cleared: unused Joe Bauman duplicate; CH6822 Miriam → Mariam Tadesse |
| L | Select Performer, Add Title records | L6823, L6837, L6839, L6842 Miriam → Mariam Tadesse |
| O | Include file name including extension | O6647:O6652 set to verified Clip #17.mov through Clip #22.mov |
| AW | Copy/paste file link | AW6647:AW6652 replaced folder URLs with exact files, preserving resource keys |
| AT | Rights Management | AT6705 set to Eric’s explicit author/feature/workshop contract category |

The source row/event/title mapping and file name, video MIME type and parent folder were verified for the six Donte files. `ericGeneratedFilenames_` then read O/AW from those source rows and populated only the six corresponding script-owned E4526:E4531 filename outputs. The F4526:F4531 video links updated through the existing source formulas. The general form-submit handler and broad filename backfill were not run.

## Results and limits

- Donte: six reported folder-to-file links repaired; all six filename inputs and generated outputs agree. Source and derived links were independently read back.
- Garner: duplicate selectable entry removed. The registry has exactly one Garner and no Krystal Garner entry. Ten historical Krystal Garner title labels/content IDs remain intentionally intact pending alias migration.
- Joel: duplicate selectable Joe entry removed; exactly one canonical Joel entry remains, and the correct title is retained.
- Mariam: registry and four title source records corrected. Independent derived readback shows four Mariam and zero Miriam titles. Their existing links, empty FV/score dependencies and titles were preserved.
- Kyle: source AT6705 and derived C4566 show the reported author/feature/workshop contract category. This records Eric’s classification; it is not an independent publication-clearance determination.
- Symone: no new repair needed. Current registry has one correct entry; all six titles use Symone Johnson-Hawkins, with zero misspelled title references.

The two performer selectors each contain 838 names and exactly match the current registry; Performer List help text was refreshed. Full source and filename invariants excluding the owned changes, plus the primary formula anchors, matched the baseline. Independently read back every changed input and the six generated filename/link outputs. 51 focused tests pass; this is not a claim about the full repository suite.

## Exact rollback proof and retained backups

Before production writes, the same transaction, source-driven population and undo code was exercised in a new isolated workbook/Form. Only affected case values were seeded; the production workbook and Form were not copied. The test deliberately used different original picker entries/help and nonblank original generated filenames, proving restoration of preimages rather than rebuilding a list from the source.

[Rollback proof workbook](https://docs.google.com/spreadsheets/d/1SGDgVqsp_e9ulDrR5p29CX0bE0GsKsuLhlNHk1ndCeM/edit), operation `51493593-2cd7-4aaa-89ee-8e907f26b0d0`: VERIFIED → UNDO_INTENT → UNDONE. Full source values, title values, formulas and exact form settings matched their before hashes. The test Form was confirmed closed to responses. No production data was written by this proof.

The production audit retains INTENT with every source before/after and each original generated filename, three FORM_BEFORE snapshots (choice order, navigation, Other setting and help), FORM_TARGET/FORM_AFTER snapshots, and final verification. `undoEricReportedRepairs(operationId)` restores the recorded source inputs, generated filenames and exact original form settings without erasing history. It refuses unexpected edits or missing backups; it does not overwrite later work to force a rollback. Production undo was not run, since the verified corrections are intended to remain applied. Writes are sequential Apps Script operations; any incomplete action stays UNCERTAIN and blocks another execution.

## Still requires review

1. Donte: source row 6458 already refers to clip 17. After the six reported corrections, clip 17 has two references and clip 15 has none. A full source scan confirmed clip 16 is already linked at row 6457; earlier provisional references to a clip-16 gap were corrected before execution. Verify content/bout evidence before replacing or merging 6458.
2. Mariam: source row 6823 has a folder URL and title B****; 6842 has an actual file with the same title. Check repair versus duplicate/merge; no row/title deletion was performed.
3. Garner: establish an explicit historical alias before migrating the ten older content IDs, so FV and score joins remain connected.

## Practical roadmap: replace email as the work queue

1. Add **Report a problem** beside an event/title in Helper, pre-filling event, source row and Drive IDs. Capture issue type, expected correction and evidence. Store case/status/owner in Helper; email remains optional evidence.
2. Add intake checks for file links resolving to folders, filename/source mismatches, repeated exact media IDs within an event/version role, and approved performer aliases. Use existing maps and folder context. Similar names are suggestions, not automatic identity merges; OM/FV and horizontal/vertical roles remain distinct.
3. Give the video team a small preview and permission to approve supported source-input repairs, with scoped derived refresh, readback and retained undo history. Sam handles ambiguous identity/rights evidence or unsupported actions. Expand automation from repeated successful cases.

Implemented now: bounded executor, separate audit, exact rollback support/proof and owner report. Still roadmap: general employee reporting/assignment, reusable detectors and video-team repair permissions. No email replies, notifications or recurring automation were created.

## Deployment integration

Preserve the current Helper project. Add EricRepairs.gs, EricRepairTransaction.gs and EricRepairRollbackTest.gs and the explicit POST cases in WeaverBridge.gs. In existing doGet, before the footage route:

```js
if (page === 'reported-repairs') {
  weaverRequireOwner_();
  return HtmlService.createHtmlOutput(ericRepairPage_()).setTitle('Reported footage repairs');
}
```

Test/source edits never change the normal employee deployment version automatically. The rollback harness rejects the production destination; the live executor validates the production ID and exact source addresses.
