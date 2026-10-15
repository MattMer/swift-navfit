# r108 — RSCA table and current workspace tools

Implemented the 14 October feedback in separate native Access blank and synthetic demo candidates. Existing workspaces and previous candidates remain intact. No GitHub push or publication was requested.

- Reporting senior > RSCA history now provides one row per paygrade, E1–E9, W1–W5 and O1–O10. Columns are RSCA, previous graded trait count, and date verified/pulled from BOL, with calendar selectors. Unused rows can remain blank.
- Save RSCA table validates every edited row, then appends all observations atomically. Failure retains pending cells and rolls back earlier rows. Section changes preserve pending entries. Prior observations remain available under View prior baselines; existing group baseline selections are not replaced implicitly.
- Existing exact trait sums survive date-only updates. Changing an average/count uses the existing estimated PSG calculation rather than deriving a supposedly exact total from a rounded average.
- Resources > Review route templates opens a dedicated template page, lists ordered reviewer steps and creates reusable routes through the existing service. It no longer opens the old Summary Groups workbench.
- Current group Members/Comparison, Defaults and Review/History pages provide the reporting senior caption, Reporting senior / PSG RSCA and Export group to NAVFIT98A. Tools receive the current group ID explicitly and drain queued comparison saves. The calculator can open the selected reporting senior record.
- Native group export continues to use its existing saved roster preview, validation and export safeguards. This change makes that workflow accessible; it adds no direct printing engine or group-return reconciliation.

App remains 0.25.0-dev/schema 25. One necessary capacity correction expands RSCABaseline.Paygrade from TEXT(2) to TEXT(5), matching other paygrade storage, so O10 fits. The additive build pass applies this idempotently; no existing rows or identifiers are removed. New blank creation uses the corrected width. This is a storage-capacity repair, not a new data model.

Validation: 116 new native resource assertions, 39 existing reporting-senior/PSG assertions, 61 live member assertions and 198 recursive-defaults assertions passed (414 total). Both candidates compile and startup/blank-data verification is recorded separately. Physical keyboard, screen reader and display/DPI acceptance remain manual; no computer-use automation was used.

Artifacts:

- `output/ux/Swift_NAVFIT_UI_Demo_2026-10-14-r108.accdb`
- `output/ux/Swift_NAVFIT_UI_Blank_2026-10-14-r108.accdb`
- `planning/RESOURCE_TOOLS_R108_VERIFICATION.json`

Source/build: modUXSeniorTable, modUXResourceTools, apply-ux-resource-tools.ps1 and build-ux-resource-tools.ps1. Scratch-only tests: ResourceToolsHarness and test-ux-resource-tools.ps1. The normal build includes the new final UI pass.
