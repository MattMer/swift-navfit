# Startup correction — r87

r86 lacked AutoExec because incremental packaging started from a form-generation candidate that had stopped before startup macro installation. The corrected r87 copies restore AutoExec and the incremental/demo builders now require it. Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r87.accdb`; matching blank: `output/ux/Swift_NAVFIT_UI_Blank_2026-10-06-r87.accdb`. r86 files are preserved, and the demo repair copies the current r86 data rather than reseeding it.

The repaired demo was reopened normally on a disposable copy and automatically loaded frmHome without explicitly invoking Startup or AutoExec. A reusable `scripts/test-ux-startup.ps1` guards both demo Overview and blank first-run startup. Existing r86 feature verification below remains valid; this correction restores startup wiring.

---

# Admin and summary-group feedback — 6 October 2026

## Implemented changes

- Report editor: a combined **Admin data** section, canonical full names instead of duplicate editable name parts, and the secondary physical-readiness result hidden. Existing hidden values remain in the editor buffer and export services.
- Block 28: **Command employment and command achievements**, mapped to NAVFIT `Reports.Achievements`. This requires additive schema **24**: nullable long-text columns on Report and RecoveryDraft, a 79th revision field, recovery, restore, duplication, shared defaults and native exchange support. Older revision snapshots retain their original field set. Older NAVFIT imports can display/export Block 28 from preserved report provenance; explicit clearing takes precedence without deleting that provenance.
- Direct group Defaults and Comparison entry points. Defaults are labelled **Defaults for new reports**; bulk changes are labelled **Change blocks on existing reports**. They retain distinct backend behavior: starting values versus guarded edits to saved reports.
- Double-click member names in the comparison, or rows in the group tools' member list, to open the report.
- Comparison cells save through the normal guarded revision transaction. Each successful save captures a fresh comparison baseline; stale saved reports remain guarded. Prior comparison history is retained.
- Reviewed-proposal application: explicit DAO database ownership in meeting and approval services; native proposal application regression passes.
- Trait guidance height follows the complete trait list.
- Active contributor selector in My Work.
- Saved snippets now seed the 20 cached attributed excerpts and original, explicitly labelled accomplishment templates for all 72 rate/source categories. Source collections remain separate from locally stored snippets and have a direct entry point. Existing snippets are retained and seeding is idempotent. Templates require replacement of placeholders with verified accomplishments.

Evidence before the final candidate: 120 focused native assertions; 16 career/table assertions; 12 member-grid/return assertions; 46 major workflow assertions; 507 NAVFIT adapter assertions; 410 exchange assertions; 99 exchange guards. Tests use disposable copies and synthetic data. No physical-display, screen-reader or native NAVFIT print acceptance is claimed.

## Final candidate — r86

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r86.accdb` to test synthetic data. New-workspace/migration template: `output/ux/Swift_NAVFIT_UI_Blank_2026-10-06-r86.accdb`. Both are 0.25.0-dev / schema 24; prior workspaces and published packages are preserved.

- 142 focused native admin, Block 28, immediate comparison save/approval reopening, hidden directory controls and rate-filtered snippet assertions passed on r86.
- 2,113 responsive geometry, focus/caret and independent identity assertions passed on r86.
- The exact populated r86 demo passed 36 native journey assertions on a disposable copy; its source hash was unchanged.
- Read-only final artifact checks: blank user tables empty; demo has 9 reports, 3 summary groups, 8 prior evaluations and 93 saved snippets (20 cached attributed excerpts, 72 original rate/category starter templates, 1 synthetic example).
- Recovery/guarded exchanges and schema transfer are verified below. Approval decisions are tested through an explicit Yes/No test input; physical interaction with the prompt remains a manual acceptance item.

Still outstanding: physical display/DPI acceptance across the target resolutions/scales, keyboard/screen-reader acceptance of all native journeys, and current NAVFIT98A application/printed-output proof. Starter templates are not a full collection of NavyWriter text; placeholders require real accomplishments.

The published GitHub 0.24.0 preview remains unchanged. Development candidates are 0.25.0-dev / schema 24.

## Additional native verification

r85: 138 focused assertions; 109 surface assertions (r84); recovery reopened in a new Access session with 12 preparation and 23 recovery assertions, active-text timer capture and Block 28 recovery. NAVFIT: 507 adapter, 410 exchange and 99 stale/hash/identity/lock/undo assertions. My Work: 32. Global defaults: 26. Group correctness: 39. Source catalog: 29. Synthetic PDF extraction/mapping/merge/undo: 256 including Block 28 for EVAL, CHIEFEVAL and FITREP.

Migration evidence: `build/tests/upgrade-20d07891-0025-4f0e-9484-4ca5e27644c0/verification.json`. Populated schema 20 and schema 23 transferred to 24; exact original rows/fields and reopened destinations verified; source/template hashes unchanged. Existing-target overwrite, unknown source tables and populated migration templates rejected.

Full form generation exposed a reporting-senior directory dependency on hidden name-part controls. r86 restores those hidden adapter controls in both incremental and full build paths. The diagnostic text now reports 75 mappings and existing group export capability.

Prototype JavaScript syntax checks and Git whitespace checks pass. New PDF fixtures were generated separately in `build/tests/pdf-block28-fixtures`; original templates/established fixtures were preserved.
