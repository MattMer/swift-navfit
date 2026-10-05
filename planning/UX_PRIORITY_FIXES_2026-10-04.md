# Native workflow priority fixes — 2026-10-04

Current source remains 0.22.0-ux-preview, schema 21. No schema changes or migration are required for these fixes; use a new build rather than replacing an existing workspace with a seeded demo.

## Behavior

1. **Close protection and save by default.** Native editor Unload and custom navigation share a close policy. Changed reports save automatically without the leave popup, per the user's updated preference. Save draft remains available. A failed save cancels normal document closing and retains pending values and inline errors. Access can force application shutdown despite a form's cancellation; that path immediately checkpoints invalid raw edits for deliberate recovery. Valid edits save before forced shutdown. Backup's explicitly approved checkpoint-and-close path retains recovery data without committing invalid values or prompting again.
2. **Current bulk previews only.** Apply starts disabled. Typing, completed field edits or changed reasons clear the preview. Apply verifies reason and group/report stamps, and refuses changed targets or saved baselines. Reopening bulk tools clears the previous batch. Successful apply retains the reviewed rows as results, disables repeat application and refreshes the target stamp. Existing atomic commits and guarded undo remain unchanged.
3. **Report-specific review context.** Review resets unrelated group/run/snapshot/batch/proposal selections, selects only the current report, filters routes to that report, and displays its member/period/saved revision. Group selectors and group-only modes are hidden. Returning to a group restores its normal selectors and navigation. Review's saved-version decision remains separate from automatic save on leaving; using the saved version retains pending editor values.

## Build and automation details

Access did not reliably retain class-module changes made during initial form generation. The builder now reopens persisted forms, creates the editor Unload procedure, compiles, reopens again and verifies the handler code exists. Behavioral tests additionally execute the actual Unload event.

Code-behind forms are inspected through typed VBA form/control references rather than unreliable late-bound PowerShell control access. Existing persistence fixtures explicitly retain their abrupt-close semantics and checkpoints; the workflow harness tests real user exits. Preview generation now refuses success if workflow checks fail.

## Verified artifacts

- Blank tested template: `build/ux-priority-fixes-verified.accdb`, SHA256 `437352AC24FC4C2A3F6817AA4D6AF19706BD941A8DA8FDDE86EB8DA650A370BC`.
- Seeded user preview: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r14.accdb`, SHA256 `BBF8ED1E92920AA110732B134FCDD8E15287E54C3741F35ACB6481345C9E654D`.
- Preview contains nine synthetic reports, three groups, two contributors, a reusable reporting senior, three RSCA baselines and review/approval/archive/NOB/recovery examples. Seeding did not change the blank template.

## Verification

- All 24 foundation suites passed during implementation, including every field family, traits, review, provenance, revision comparison, recovery, imports, major workflows and limited rules.
- After the user's final automatic-save preference, the final template passed smoke, recovery, backup, 425 shell bounds checks, 38 reporting-senior checks, six workbench assertions and 76 visible-control bounds checks.
- Preview generation passed 108 native surface/workspace assertions, 50 workflow assertions, both forced-shutdown scenarios, **36 populated-demo assertions**, blocked-content startup and enabled-content startup. The final caption update raised the workflow coverage to **51 assertions**, rerun successfully with the forced-shutdown and demo checks.
- Workflow cases include active textbox text before the first timer, automatic close without a prompt, failed save retaining values, explicit cancel/discard, current preview requirements, reason changes, membership changes, repeated apply, review outside a prior group, retained pending review edits and restored group navigation.
- Forced-shutdown tests use separate processes/databases and read saved/recovery values after Access exits. Valid edits were committed; invalid SSN and official-name pending values were retained in RecoveryDraft without committing the invalid entry.

Final check logs are in `build/tests/priority-final-*.log`; foundation results are in `build/tests/foundations-suite-results.json`. These generated files and databases are excluded from Git. The earlier full-suite run preceded the final automatic-save default; focused lifecycle and integration checks above were rerun afterward.

Existing user workspaces and the published 0.21.0 binary were unchanged. This is a local development preview, not a new GitHub binary release. Physical layout, Windows scaling, screen-reader behavior and Runtime/x86 acceptance remain pending; these fixes do not finish the broader redesign.

The final report-context caption uses the report revision number rather than an internal revision ID. Both template and demo received this caption update; the 51 workflow assertions, forced-shutdown cases and 36 demo assertions passed again afterward.
