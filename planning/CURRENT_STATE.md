## Published r104 narrative editor prerelease (10 October 2026)

Fixed Courier New/plain-text comments geometry by report type; supplied EVAL PDF measured. Open output/ux/Swift_NAVFIT_UI_Demo_2026-10-10-r104.accdb for the synthetic Lorem Ipsum example; matching truly blank is Swift_NAVFIT_UI_Blank_2026-10-10-r104-clean.accdb. Existing locally populated r103 preserved. See [layout details and limits](NARRATIVE_LAYOUT_R104_2026-10-10.md). App/schema unchanged; r104 prerelease includes verified blank and synthetic demo. Older releases remain available.

## Current blank deployment — r103 (10 October 2026)

Latest blank artifact: output/ux/Swift_NAVFIT_Deployment_2026-10-10-r103-blank.accdb. App 0.25.0-dev/schema 25; includes r99–r102 efficiency/recovery work and function comments. All 48 user-data tables empty; setup startup and exact-file scratch tests verified. See [deployment notes](DEPLOYMENT_R103_2026-10-10.md). Earlier entries below are historical checkpoints.

## 2026-10-10 - larger synthetic groups measured

Created a separate 525-report scale database: 25/50/100-report groups for each report type. Measured native group opens, row navigation, queued saves and editor opens on scratch copies. Row moves 5-7 ms; saves 93-110 ms; group opening 0.64-1.09 s; report editor opening 2.44-4.01 s. Exactly 90 tested edits produced 90 revisions and zero group-history rows. Next priority: profile editor startup before optimizing. See [scale results](SCALE_R102_2026-10-10.md). No production source or schema changes.

## 2026-10-10 - r102 efficiency priorities 3-6 complete

Completed targeted editor/group refresh caching; removed duplicate hidden comparison list and adapted proposal selection to the grid; routed Members directly and simplified group tab transitions; made validation/recovery timer work event-driven. Idle calls produce zero validation scans. Saved averages refresh after commit; typing and field-exit checks/recovery remain. All 49 user tables preserved; schema 25. See [implementation and verification](EFFICIENCY_R102_2026-10-10.md). Physical display/accessibility and scale performance verification remain outstanding.

## 2026-10-10 - r101 live Members and cached display guards

Completed single-user efficiency priorities 1 and 2. Members reads current reports directly and no longer creates historical snapshots on opening or direct cell saves. Row focus uses loaded state; authoritative checks run at save. Historical proposal baselines/conflict checks and pending-save recovery remain. Three group opens add zero history rows. All 49 user tables preserved; schema 25. See [implementation and verification](LIVE_MEMBERS_R101_2026-10-10.md). Next: target editor refreshes, prove/remove legacy comparison controls, then simplify group page navigation.

## 2026-10-10 - r100 performance candidate

Completed unchanged comparison baseline reuse, input-driven live validation caching, unchanged checkpoint write suppression, joined saved-average reads, and conditional comparison ITA writes. Three unchanged group opens add 8 rather than 24 history rows; timing speedup is not established. Pending-save safeguards retained; schema 25; all 49 user tables preserved. See [performance notes](PERFORMANCE_R100_2026-10-10.md). Successful cell-save snapshot creation remains a follow-up.

## Current local candidate — r99 (10 October 2026)

Highest-priority review fix implemented: comparison save failures retain failed/later queued cells, with Retry/Discard actions and native close cancellation. Contributor changes, proposal application, group-page transitions and report creation respect pending cells. Already saved revisions remain intact; normal valid saves stay automatic. In-app version/schema help now uses constants through the canonical help generator.

Open output/ux/Swift_NAVFIT_UI_Demo_2026-10-10-r99.accdb for synthetic practice, matching blank for a new workspace, or Swift_NAVFIT_UI_TestWorkspace_2026-10-10-r99.accdb for retained r97 data. All 49 user tables are unchanged. See PENDING_CELL_RECOVERY_R99_2026-10-10.md and its verification JSON. r98 remains the published baseline; r99 has not been pushed/released. Snapshot lifecycle and change-driven editor performance are next. Pending comparison recovery is session-scoped, not crash/startup recovery.

## Current deployment baseline — r98 (10 October 2026)

The current baseline is application 0.25.0-dev/schema 25, packaged as a development prerelease with fresh blank and synthetic demo files. Use output/Swift_NAVFIT_Baseline_2026-10-10-r98/Swift_NAVFIT_Deployment_2026-10-10-r98-DEMO.accdb for testing, or the matching blank for a new workspace. Existing author-edited and r97 user workspaces are preserved. See BASELINE_R98_2026-10-10.md and BASELINE_R98_VERIFICATION.json for exact scope, checks and hashes. No new narrative-fit or returned-group feature is included; native controller proof remains partial, with full round-trip, printing, physical accessibility and migration acceptance outstanding.

## Current — r97 RUAD import / refresh (9 October 2026)

The user confirmed r96 resolves the interactive comparison save failure. r97 adds Resources → Command member directory → Import / refresh RUAD CSV: explicit identity selection, complete 35-column preview, six vacant-row skips, provisional new members, three-way roster refresh with manual conflict preservation, stale/duplicate guards and atomic apply/results. Source baseline/path/time use existing RUADData XML; schema remains 25. No absent members are automatically removed. Archive filter now says Archive status.

Open output/ux/Swift_NAVFIT_UI_TestWorkspace_2026-10-09-r97.accdb. All 49 existing tables match r96; 71 RUAD assertions, 26 save/approval assertions and 354 authored-baseline/startup/navigation assertions pass. No computer use or publication. Details: RUAD_IMPORT_R97_2026-10-09.md. Current outstanding work: REMAINING_WORK_2026-10-09.md. Prior candidate entries below are historical and their unconfirmed-fix statements have been superseded by the user's r96 confirmation.

## 9 October — r96 revision cursor fix

User trace isolated DAO 3218 at revision-pointer linking. CommitDraft now links through the existing report write cursor instead of a second SQL UPDATE; closed revision cursor is cleared before cleanup. r96 workspace preserves all 49 user tables and local diagnostics. Native compile, 26 approval/rollback/retry checks, three report-type event saves, queued-cell checks and existing-report timer save passed. User's exact interactive journey still needs verification. See REVISION_CURSOR_R96_2026-10-09.md.

## 9 October — r95 diagnostic navigation candidate

Populated groups now open their editable member table without rendering the legacy Members page first. Removed duplicate initialization; group tab changes skip the redundant list rebuild. Native compile, scratch navigation assertions and 26 save/approval assertions passed; all 49 user tables preserved. Interactive 3420 failure remains unresolved: r95 includes a local save-stage trace. See INTERACTIVE_SAVE_DIAGNOSTIC_R94_2026-10-09.md. This is a diagnostic candidate, not a published release.

# Current patch candidate — r93 (9 October 2026)

Open `output/ux/Swift_NAVFIT_UI_TestWorkspace_2026-10-09-r93.accdb` to retain the current test group/report, or the matching blank/demo. Comparison controls queue their changes; the form timer commits after Access completes its bound-cache update. Closing flushes pending cells. Native timer saves, E6 1.0 marks, row identity, coalescing, failed-edit reset and retry pass. All 49 user tables remain identical to r92. The exact user failure was not reproduced; retest is still required to confirm resolution. See COMPARISON_EVENT_SAVE_R93_2026-10-09.md. No GitHub publication.
# Current patch candidate — r92 (9 October 2026)

Retest with `output/ux/Swift_NAVFIT_UI_Demo_2026-10-09-r92.accdb` or matching blank. The debugger identified error 3218 at the revision-pointer update. The patch removes its second database write handle and uses the report transaction's existing handle for history writes too. Native shared-mode control events from newly created reports pass for all three report types. Exact user-failure resolution remains unconfirmed. See REVISION_LOCK_R92_2026-10-09.md. No GitHub publication.
# Current patch candidate — r91 (7 October 2026)

Use `output/ux/Swift_NAVFIT_UI_Demo_2026-10-07-r91.accdb` to retest the reported O3 comparison trait-save error. Matching blank is available. DAO lifetime hardening is compiled and regression-tested; the exact user failure was not reproduced in r90. All 49 user tables are identical to the current r90 source. See TRAIT_SAVE_R91_2026-10-07.md. No GitHub publication.
# Summary-group and directory feedback — r90 (7 October 2026)

Current demo: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-07-r90.accdb`; matching blank: `output/ux/Swift_NAVFIT_UI_Blank_2026-10-07-r90.accdb`. Version 0.25.0-dev/schema 25. The author's edited database and r89 remain preserved; GitHub release is unchanged.

- Completed: Add report from a summary group, choose an existing directory member or enter a new member, inherit saved group defaults, create/link atomically and return to group members after editing. Cancel creates nothing. Stale group/default/member changes and duplicate member/type/period are rejected.
- Completed: common Members, Defaults for new reports, Review, Comparison and History navigation on every group workspace. Switching Members/Comparison keeps the existing comparison context. Pending default edits cannot disappear through navigation.
- Completed: Overview contributor selector/status and Unposted notes occupy separate areas from cards, recovery actions and recent reports.
- Completed: member profiles expose all 35 RUAD columns in five compact pages. Text identifiers retain leading zeroes; profile saves retain hidden-page values and use version checks. Two synthetic complete profiles are included in the demo.

Schema 25 adds only nullable CommandMember.RUADData memo storage. Full RUAD CSV import/refresh remains pending. See GROUP_DIRECTORY_FEEDBACK_2026-10-07.md for verification and remaining UX work.

## Previous r89 baseline (historical)
# Author-edited baseline — r89 (7 October 2026)

Use `output/ux/Swift_NAVFIT_UI_Demo_2026-10-07-r89.accdb` for continued UI work. Matching blank: `output/ux/Swift_NAVFIT_UI_Blank_2026-10-07-r89.accdb`. Version 0.25.0-dev/schema 24; published GitHub release remains unchanged. The supplied `Swift_NAVFIT_UI_Demo_2026-10-06-EDITED.accdb` and r88 are preserved.

## Reviewed changes

Seven forms changed: frmLibrary, frmMajorWorkbench, frmMyWork, frmUXGroups, frmUXImportChoose, frmUXMemberHistory and frmUXResources. There were 316 property/object differences, with no standard VBA module or form-class code changes. Adopted tighter report filters/actions, revised Validation/Move report/Summary Group wording, the new Column list and Archive status labels, grouped summary-group actions/defaults, the My Work contributor placement, the shorter history explanation and stronger member heading, and import/resources alignment changes.

The baseline is reproducible through `src/ui-author-baseline.json` (303 property changes and two added labels), applied by `scripts/apply-ui-author-baseline.ps1` at the end of the current full and incremental build path. Runtime My Work layout preserves the author's contributor placement at normal widths, while retaining the compact-width layout. Raw report-list ColumnWidths are not adopted because the nine-column adapter computes them from the selected view and must keep internal IDs hidden. Some conditional workbench controls intentionally retain their runtime positions when their tool opens rather than displaying the author's off-screen design positions.

## Hidden-control cleanup

The r88 inventory contained 524 hidden controls. Traced source references, event bindings, tags, data sources and dynamic families before selecting removals. Removed 50 additional obsolete controls from the author's copy; including eight authored deletions, the candidate has 58 fewer obsolete controls than r88. These include old Home cards/actions, duplicate editor labels/detail buttons and old bulk page buttons. The exact allowlist and reasons are in the baseline JSON.

Retained unbound report/bulk buffers, IDs/stamps, conditional section/error controls, responsive navigation, trait labels addressed by constructed names, reset/inheritance buttons and reporting-senior adapter controls. Restored only `frmLibrary.cboDestination` as a hidden compatibility buffer: older folder functions still reference it. No tables, forms, modules or backend services were removed.

## Verification

Compiled both candidates. All 49 user tables retain identical field schemas, row counts and row-value hashes compared with the author's returned database, including its additional report, group, revisions and note. Both originals' hashes stayed unchanged during inspection.

Passed 2,684 assertions: 142 admin/comparison, 2,107 responsive/focus/identity, 22 group organization, 32 My Work, 355 author-property/cleanup/startup/populated-navigation and 26 approved-trait retry checks. Blank normal startup displays Setup; the demo startup check displays Overview. Native visual spot-checks covered Reports, Summary groups and My Work. This does not establish screen-reader or physical DPI acceptance.

The supplied file carries Zone.Identifier. It was preserved on the original; removed only from the newly generated, reviewed local build. Access startup protection remains intact. No global trust or security setting changed.

Two visual follow-ups remain: report columns truncate some headings/names at the observed window width; the Summary groups archive-state filter says Group status even though the workflow status selector is separate below. Recorded for the next UX pass rather than changing the author's design beyond the requested cleanup.

## RUAD

Received and inspected the synthetic sample. It has 35 columns, 66 rows, 60 named members and six vacancies, with no malformed rows. Mapping details and identity/refresh behavior are documented in `RUAD_MAPPING_2026-10-07.md`; counts in `RUAD_SAMPLE_ANALYSIS_2026-10-07.json`. Earlier sample-blocked notes are historical. RUAD import and refresh implementation is the next remaining feature; it is not included in r89.

Evidence: UI_AUTHOR_REVIEW_VERIFICATION_2026-10-07.json and UI_AUTHOR_DATA_PRESERVATION_2026-10-07.json. Use the current full builder or build-ux-admin-feedback.ps1; historical incremental recipes do not represent the new baseline.

---

# r88 comparison retry and native dates — 6 October 2026

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r88.accdb`; matching blank is `Swift_NAVFIT_UI_Blank_2026-10-06-r88.accdb`. Demo copied from the current r87 workspace without reseeding; existing practice records preserved. Version 0.25.0-dev/schema 24; GitHub release unchanged.

Comparison edits now isolate the temporary row database in its own DAO workspace, capture the expected revision before the approval prompt, guard against reentrant edit callbacks, retain owned database references in dependent operations, and restore the row after failed saves. Errors include the failing save stage and error number. The misplaced `stage` diagnostic in UXGridCurrent was removed before this build.

Native calendar selectors are enabled for report dates and applicable shared/default, prior-evaluation, review-due and reporting-senior date fields. Calendar dates normalize to yyyy-mm-dd; invalid typed entries remain recoverable and counseling status text remains supported.

Verification: compiled blank and demo; 128 calendar assertions (all three report types, raw invalid recovery and counseling statuses); 26 approved-trait rollback/retry/repeated-save/reopen assertions; 142 admin/comparison assertions; 46 major workflow, 39 summary-group correctness and 26 global-default assertions. Both packaged copies open their normal startup forms. Visual check on a disposable r88 demo confirmed the Date Reported calendar icon and calendar popup in native Access.

Limit: the original object-invalid error was not reproduced in clean r87 native approved-report tests. Even the older build passed the corrected synthetic failure/retry test. These changes harden transaction and object lifetimes; they do not establish a reproduced root cause. If it recurs, the new stage/error diagnostic is needed. No new physical scaling, screen-reader, NAVFIT printing or GitHub publication claim.

---

# Startup correction — r87

r86 lacked AutoExec because incremental packaging started from a form-generation candidate that had stopped before startup macro installation. The corrected r87 copies restore AutoExec and the incremental/demo builders now require it. Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r87.accdb`; matching blank: `output/ux/Swift_NAVFIT_UI_Blank_2026-10-06-r87.accdb`. r86 files are preserved, and the demo repair copies the current r86 data rather than reseeding it.

The repaired demo was reopened normally on a disposable copy and automatically loaded frmHome without explicitly invoking Startup or AutoExec. A reusable `scripts/test-ux-startup.ps1` guards both demo Overview and blank first-run startup. Existing r86 feature verification below remains valid; this correction restores startup wiring.

---

# Latest development candidate — r86 (6 October 2026)

Test with `output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r86.accdb`.
Blank template: `output/ux/Swift_NAVFIT_UI_Blank_2026-10-06-r86.accdb`.
Version 0.25.0-dev, schema 24. GitHub 0.24.0-preview remains the published release.

Admin data combines member/report and senior sections, with Block 28 and hidden redundant fields. Comparison cells save immediately, with approval reopening; Defaults/Comparison open directly; My Work has its contributor selector; all rate/source categories have saved starter snippets. Hidden directory adapter controls are retained.

Verification: 142 focused assertions, 2,113 responsive assertions, 36 exact-demo assertions; recovery, native exchange, PDF Block 28 and schema-20/23 transfer checks pass. See `ADMIN_COMPARISON_FEEDBACK_2026-10-06.md` and `ADMIN_COMPARISON_VERIFICATION_2026-10-06.json` for candidate-specific evidence and limitations. Physical scaling/screen-reader acceptance and NAVFIT98A GUI/print proof remain outstanding.

---

# Deployment prerelease — 0.24.0-preview (6 October 2026)

Package: `output/Swift_NAVFIT_0.24.0-preview.zip`, schema 23.
Open `Swift_NAVFIT_0.24.0-preview-DEMO.accdb` for synthetic testing or use
`Swift_NAVFIT_0.24.0-preview-blank.accdb` for a new workspace/migration template.
The release is based on r75 with updated metadata and release-check fixes: schema-23 backup/restore, and review stamps that remain current after local approval locks. Bulk edit-lock guards remain intact; prior workspaces/builds remain.
Exact packaged-byte verification: 27 suites, 72 matching source modules, empty blank user tables, and schema-20/21/22/23 transfer acceptance. Hashes and logs are in the package MANIFEST.json; repository evidence is DEPLOYMENT_0.24.0_VERIFICATION.json.
See RELEASE_0.24.0_PREVIEW.md for changes since 0.23.0 and migration instructions.
The older milestones below remain historical evidence.

---
# Latest development increment â€” r75 (6 October 2026)

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r75.accdb` for synthetic testing.
Blank candidate: `build/ux-r75-policy-blank.accdb`; schema 23, published release unchanged.

Narrative sources now include 72 rate/source collections and 20 excerpts, with
rate/search filters in the catalog and saved snippets, plus manual rate assignment.
Live report and summary-group checks now cover counseling, required administrative
fields, career recommendations, partial NOB, adverse justification/referral and
extended Regular periods. Counseling NOT REQ/NOT PERF survives save, PDF import
and NAVFIT conversion. Every finding retains its policy locator and excerpt.

Evidence: 37 citation/page checks, 29 narrative UI assertions, 102 live-validation
assertions, 170 service-detail and 49 readiness checks, 501 adapter / 410 exchange /
98 guard assertions, 250 existing PDF checks and 12 new counseling PDF checks.
See NARRATIVE_POLICY_REVIEW_2026-10-06.md for rule coverage, known gaps and
prioritized research recommendations. Final r75: 72 installed modules match source;
36 populated-demo, 30 guide and 29 populated catalog assertions pass. Transfer
preserves all original rows/fields and the r68 source hash. Automated responsive
checks on r74 passed 1,972 assertions; this does not establish physical DPI acceptance.
Final demo SHA-256: `8B1BED68A250AD6ECDB3CCF32B2E3D48C6C23E73474254141B89D48ADA44A758`.
No Navy-wide compliance certification,
printed-fit claim, release publication or Git push is implied.

---
# Latest development increment - r68 (6 October 2026)

Populated demo: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r68.accdb`.
Blank: `build/ux-r68-identity-blank.accdb`. Development only, schema 23;
published schema-21 release remains unchanged.

Adds the command directory, automatic report identity links, manual prior
evaluation history, prior ITA/promotion in group member grids, explicit identity
confirmation, and previewed atomic report creation for selected directory members.
The demo contains nine synthetic reports, three groups, nine directory profiles and eight synthetic prior evaluations.

Recent verification across r66-r68: 20 batch, 21 directory/history UI, 21 identity/history service,
14 member-grid and 49 workflow assertions. The workflow suite requires the blank
fixture; the populated demo intentionally includes a recovery checkpoint.
Schema 22-to-23 and 23-to-23 transfers preserve every original row/field, source
hashes and existing identities. All 72 installed source modules match; 244 native group import checks and 36 populated-demo checks pass. Notes and earlier feature checks are recorded below.

Outstanding: verified RUAD mapping/import (sample path pending), exact narrative
print-fit behavior, physical display/scaling and screen-reader checks, and native
NAVFIT98A UI round-trip acceptance. No new deployment has been published.

---
# Latest development increment - r56 (6 October 2026)

Populated demo: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r56.accdb`.
Blank: `build/ux-r56-notes-blank.accdb`. Development only, schema 22; published
schema-21 release remains unchanged. This adds attributed note entries,
full-text viewing, unposted-note checkpoints and recovery from Overview.
Notes remain separate from NAVPERS content and report revisions. Legacy notes
retain their original saved field and become entries with unknown attribution.

Verified: 68 installed source modules; 1972 layout/focus checks; 49 workflow
checks; report and raw-note shutdown recovery; 19 note services and 18 note UI
checks; schema-21 -> 22 and schema-22 -> 22 data transfers. See
SCHEMA22_REPORT_NOTES.md and MANUAL_FEEDBACK_PROGRESS_2026-10-06.md.
Manual directory and prior-evaluation history work is starting; the RUAD sample
path is still pending. Physical display/accessibility and native print-fit
acceptance remain outstanding.

---
# Latest development increment - r49 (6 October 2026)

Populated demo: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r49.accdb`.
Blank candidate: `build/ux-r49-blank.accdb`. Development builds only; published
0.23.0-preview remains unchanged. Schema remains 21.

Manual-feedback plan implementation now includes contributor selection,
startup pane hiding, snippet return, failed-save undo, summary-group organization,
editable member comparison/career recommendations, ITA/promotion/group columns,
group workflow, global/inherited defaults, report-header review actions,
automatic export filenames, a source-linked NavyWriter starter catalog and
selected native-group/report import. Working assumptions and per-item checks
are in MANUAL_FEEDBACK_PROGRESS_2026-10-06.md.

Selected native groups split by report kind/worn paygrade, preview all supported
report values and shared folder defaults, and import atomically with native
provenance. Duplicate imports require returned-file reconciliation. Empty folders
and unresolved/mismatched rank/type require separate source correction. Native
GUI validation/printing is not certified by the data round-trip tests.

Still outstanding: RUAD/directory/history and prior-eval columns (sample path
requested), attributed note entries/additive schema, narrative print-fit proof,
and physical display/keyboard/screen-reader acceptance. No release published.

---
# Latest development increment - r32

PR #15 is merged into main. Summary-group work continues on
`codex/summary-group-correctness`. The populated r32 demo is
`output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r32.accdb`; blank is
`build/ux-r32-check-blank.accdb`. Live group capacity is withheld for mixed/unresolved
observed keys; NOB is excluded from numerator and denominator. See
SUMMARY_GROUP_CORRECTNESS_R31.md and SUMMARY_GROUP_POLICY_R32.md for scope and
remaining work. r32 adds recorded UIC/category context and deliberate local
roster completeness, stale review detection and membership/bulk consequence
previews. Verified: 39 group correctness, 90 shared blocks/comparison, 41 group
lifecycle and 59 live-validation assertions; 58 installed modules match source,
native review/return bindings and compile pass on the populated demo.
Also passed: 1,804 responsive layout/focus/identity assertions, 36 populated-demo
assertions and 20 fingerprinted policy-excerpt checks. Manual display and
screen-reader acceptance remains outstanding.
The published 0.23.0-preview deployment below is preserved.

# Deployment version - 0.23.0-preview, schema 21

The deployment package is `output/Swift_NAVFIT_0.23.0-preview.zip`, containing
separate `Swift_NAVFIT_0.23.0-preview-blank.accdb` and
`Swift_NAVFIT_0.23.0-preview-DEMO.accdb`, migration tooling and guides. Open the
DEMO for synthetic examples. Existing workspaces need source-preserving transfer
into the blank build. See RELEASE_0.23.0_PREVIEW.md and the package MANIFEST.json
for verification and remaining acceptance work. Prior milestone filenames and
versions below are retained as historical evidence.

# Latest workspace update - r30, 2026-10-05

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r30.accdb` for the populated demo.
Blank: `build/ux-r30-blank.accdb`. Adaptive workspace layout now skips unchanged
sizes, and context refresh sizes only the active workspace. Redundant list
requeries were removed. Both builds compile; all 57 demo modules match source,
native event guards pass, and 1801 responsive assertions pass. See
WORKSPACE_PERFORMANCE_R30.md. Siltstone colors and existing feature scope remain.

# Latest presentation update - r29, 2026-10-05

The Siltstone palette is applied across all 51 native forms (1,717 controls).
Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r29.accdb` for the populated demo.
The separate blank is `build/ux-r29-final-blank.accdb`. Both compile; the demo's
57 installed modules match source and native event guards remain intact.
Version and schema remain 0.22.0-ux-preview / 21. See SILTSTONE_THEME_R29.md for
roles and verification limits. All r28 data and feature scope remain applicable.

# Current state - 2026-10-05

Swift NAVFIT is a native Microsoft Access application for local EVAL, CHIEFEVAL
and FITREP drafting. The current source milestone is r28, version
`0.22.0-ux-preview`, schema 21. The redesign is a development preview under
user review. The prior distribution checkpoint remains 0.21.0, schema 20.

## Latest runnable build: r28

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r28.accdb` for the populated
group-export preview. The separate blank template is `build/ux-r28-blank.accdb`.
Both compile and all 57 installed modules match current source. Summary-group
export previews saved members, computes promotion counts and produces one native
folder; exact PSG RSCA export is optional. Single-report export/returns remain.
Group returns and direct Swift printing are pending. Print exported reports using
NAVFIT98A's own engine after native validation. See
[group-export scope and checks](NAVFIT_GROUP_EXPORT_R28.md) and
[print-engine inspection](NAVFIT_PRINT_ENGINE_FINDINGS.md). Group native UI/print
acceptance remains unverified; prior r27 actual per-report round trips remain
the native evidence.

## Previous runnable build: r27

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r27-reviewed.accdb` for nine synthetic
reports, three groups, review work, reporting seniors/RSCA and recovery examples.
The separate blank template is `build/ux-r27-blank.accdb`.
Both artifacts compile and all 57 installed modules match current source. The
final demo passes synthetic-data and normal/enabled-content startup checks.
Exact hashes and native proof receipts are in
[r27 verification manifest](NAVFIT_R27_VERIFICATION.json).

All 74 supported saved report values plus type are mapped for EVAL, CHIEFEVAL
and FITREP. All 128 native report columns and native folder/summary content are
retained. Full supported-value import, single-report export and sealed-baseline
three-way returned-file reconciliation are implemented. Actual NAVFIT98A
2.2.0.332 edits were returned to Swift for each type; all 75 decoded values and
128 re-exported native report columns matched. Empty/null and long Unicode
preservation fixes are included. See [scope and proof](NAVFIT_EXCHANGE_R27.md).

Complete summary-group/batch export, legacy non-date counseling strings, changed
cross-file identities and other NAVFIT versions remain pending. Printed fit and
complete Navy readiness certification remain unverified. No GitHub release was
published; previous builds and populated workspaces remain preserved.

## Previous milestone: r26

The user's clarified priority #4 is implemented: expanded structured PDF import
catalogues (40 EVAL, 51 CHIEFEVAL, 36 FITREP saved fields), reviewed existing-draft
merges, stale/lock/recovery guards, source receipts and compensating undo. The
native extraction check and portable target-computer kit verify real XFA reads.
Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r26.accdb` for synthetic exploration;
`build/ux-r26-final.accdb` is the separate blank template. See
[PDF scope, evidence and workstation check](PDF_IMPORT_R26.md). Full field coverage,
other target-machine results, Runtime/x86 and physical/accessibility acceptance
remain outstanding. Earlier workspaces/builds remain preserved.

## Previous milestone: r25

The r25 rebuild adds live report diagnostics and an all-member summary-group issue list, with instruction, enclosure, page, paragraph/table and excerpts for verified Navy findings. See [scope and verification](UX_LIVE_VALIDATION_2026-10-04.md). Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r25.accdb` for synthetic exploration; `build/ux-r25-candidate.accdb` is the separate blank template. Full physical-display/accessibility acceptance and broader policy exceptions remain pending.

## Previous milestone: r24

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r24.accdb` for synthetic exploration. The separate empty workspace is `build/ux-r24-blank.accdb`. Version remains 0.22.0-ux-preview, schema 21; earlier builds and existing workspaces are preserved.

The three native-review priorities are implemented: safe report-list capture/reopening; one My Work inbox for assigned reviews, current route steps, ready local approvals and related unresolved threads; consistent heading/filter spacing, visible Save and lock feedback, readable list headings and supporting-screen navigation. Current contributors see actionable route steps rather than future steps. Stale assignment decisions are disabled. Opening work selects the actual route/thread. Resolution removes threads from the pending queue; history remains available.

Focused checks passed: 145 review/return assertions (including create, save, single-result search and reopening), 32 unified inbox assertions, 38 routing assertions and 28 comments assertions. All 50 installed standard modules match source and compile in both final artifacts; all 42 blank user-data tables are empty. Full preview verification and native visual results are recorded in `planning/UX_MY_WORK_2026-10-04.md`.

This remains a development preview. Full keyboard, screen-reader, actual Windows scaling, Runtime/x86 and release acceptance remain outstanding. No replacement GitHub release was published. The r22 candidate ZIP below remains an earlier packaged checkpoint.

## Latest UX work

- Persistent Overview, Reports, Summary groups, My Work, Resources and
  Settings & Help navigation, with command and contributor context.
- Section-based report editing with all 78 saved fields retained. Official
  captions vary by report type; structured name parts remain independent of
  stored full names. Normal editing uses SSN; legacy DoDID data is retained.
- Resizable report workspace with compact report block rows, adjacent checkbox
  captions, content-sized fields and closer label/input spacing throughout.
  Empty validation space is hidden; actual errors expand rows. Trait marks use
  compact rows with adjacent guidance. Review and History have focused layouts.
- Explicit Save draft, pending-change indicators, atomic revisions, recovery
  checkpoints, inline validation and a clickable error summary.
- Compact report layout replaces both navigation rails with selectors in narrow
  or short windows; larger windows restore the rails. Active text and caret
  survive reflow. Short windows retain vertical access to the action area.
  Group, comparison, senior and My Work destinations share viewport navigation.
- Report identities prefer the official block-1 name, with a labeled local name
  in the editor when different. Empty official names fall back to the local name,
  then Unnamed member. Saved names and name parts remain independent.
- Changed reports save automatically when leaving or closing, without a leave
  confirmation. Failed saves cancel normal report closing and preserve pending
  values. Forced Access shutdown retains raw recovery values if saving fails.
- Bulk changes require a current preview; typing, reason changes and changed
  group/report baselines invalidate Apply. Report Review resets unrelated group
  context and restricts the report and route selectors to the current report.
- Shared report blocks and group defaults use the same compact, official-block
  sections as report editing, with explicit mixed/blank/staged/clear states,
  full-value preview, atomic Apply and concurrent-default guards.
- Comparison has editable native trait cells, pending ITA and group averages,
  saved-order reload, auditable proposal replacement, before/after preview and
  stale/archive/approval guards. Temporary grid data detaches on native close.
- Summary groups support rename, reversible archive/recover, Active/Archived/All
  filters, status and retained report counts. Archived group workspaces explain
  recovery requirements; report data and group history remain intact.
- Summary-group list with visible current report counts; focused member,
  defaults, review, comparison and history views. Trait comparison captures
  its baseline automatically. Reporting-senior records and rank-specific
  historical RSCA/count baselines support group PSG RSCA calculations.
- Calculation details distinguish pending editor values from saved revisions,
  show sums/counts/exclusions and manual historical baseline provenance. PSG
  RSCA exposes exact/estimated, partial and unavailable reasons on initial load.
- Saved-draft checklist uses official captions, omits legacy/inapplicable inputs,
  and opens the relevant report section while preserving pending edits.
  Current guidance distinguishes incompleteness from limited Navy checks.
- A native workspace guide in Settings & Help covers current workflows,
  backed by a matching written guide and a complete feature-destination table.
- Contextual imports follow Choose/Preview/Apply/Results, with official mapping
  captions, full values, stale-preview guards and created-report navigation.
  SSN is the normal PDF identifier; legacy DoDID mapping requires opt-in.
- Guarded startup: blocked Access content opens an instruction screen instead
  of invoking unavailable VBA. Office security settings are not changed by builds.

Existing drafting, library/folder operations, snippets, revision history,
local review/approval, bulk transactions and undo, imports and provenance,
backup/restore and migration services remain in the project. See the
[feature inventory](../README.md#current-features) and
[major foundations guide](MAJOR_FOUNDATIONS_GUIDE.md) for their supported scope.
Local approval is not an official signature; limited Navy checks are not a
complete official validation engine.

## Earlier packaged r22 preview and verification

The earlier r22 local review artifact is
`output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r22.accdb`. It includes nine synthetic
reports, three groups, two contributors, a reporting senior with three RSCA
baselines, and review, approval, NOB, archive and recovery examples.
Earlier user workspaces and previews are preserved.

The r22 preview passes 30 native-guide/documentation consistency checks,
43 guided-import state/preview/result assertions, 28 native PDF assertions
and 12 read-only import-review checks. NAVFIT commit checks passed quarantine,
duplicates, rollback and all 2,472 retained provenance values. It also passes
23 calculation/checklist, 41 group-lifecycle, 90 shared-block/default/comparison,
1,790 responsive/focus/identity and 108 native editor-surface assertions.
It also passes 51 close/bulk/review assertions, two forced-shutdown scenarios,
36 populated-demo assertions and both startup modes.

The r20 checkpoint additionally passed 42 senior/RSCA, 54 integrated tools,
53 administration, 75 traits and 169 catalogue-fidelity checks. Those are prior
service evidence, not newly run r22 suites. All 24 foundation suites last passed
at r15; this is historical evidence. Exact packaged candidate tests are listed
in its MANIFEST.json, separate from these preview checks.

The blank is `build/ux-r22-blank.accdb`. The preview builder preserves that
source while seeding a separate demo with nine reports and three groups.
The candidate packager compares all source standard modules against installed
exports after normal VBA compiler formatting. It independently verifies blank
user-data tables, installed source, native unload handlers and migration.
The Member parameter prompt remains fixed by ordering the comparison query
by its full expression rather than its alias.

See [guided import changes](UX_IMPORT_FLOW_2026-10-04.md),
[workspace guide](UX_WORKSPACE_GUIDE.md),
[calculation and checklist changes](UX_CALCULATIONS_CHECKLIST_2026-10-04.md),
[group lifecycle evidence](UX_GROUP_LIFECYCLE_2026-10-04.md),
[shared blocks and comparison evidence](UX_BLOCK_TOOLS_2026-10-04.md),
[responsive layout and identity evidence](UX_RESPONSIVE_IDENTITY_2026-10-04.md)
and [earlier priority fixes](UX_PRIORITY_FIXES_2026-10-04.md).

Build-time SHA256:
- Demo: `1F172ABDF185DC939EA712A6E78246D3FAF2ED68E96DB69EDD29CB0D1F1AC7F1`
- Blank: `F958BA5A1A3488B05EADBFC1583DD2C7BA3DA49B243626EEB2937E6CF50DE4E1`

Opening or editing an Access file may change its hash. Full native DPI and
assistive-technology acceptance still remains.

This repository tracks source, build/test scripts, the browser design prototype
and documentation. Access databases, local reference PDFs and user workspaces
are excluded from Git. Local demos are separate from published binary releases.
For a new synthetic review build, run:

```powershell
powershell.exe -NoProfile -File scripts/build-preview.ps1
```

Use `scripts/build.ps1` for a separate blank template. The browser prototype is
an interaction design reference, not a replacement application. See the
[demo guide](UX_DEMO_BUILD_GUIDE.md) and
[checkpoint details](UX_NATIVE_LAYOUT_CHECKPOINT.md).

## Local candidate packaging

Use [the candidate guide](UX_CANDIDATE_GUIDE.md) for blank/demo distinctions,
source-preserving transfer and remaining acceptance work. Reproducible packaging:

```powershell
powershell.exe -NoProfile -File scripts/package-ux-candidate.ps1 `
  -TemplatePath build/ux-r22-blank.accdb `
  -DemoPath output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r22.accdb `
  -OutputDirectory output/candidates/Swift_NAVFIT_0.22.0-ux-preview-r22
```

The completed local package is
`output/candidates/Swift_NAVFIT_0.22.0-ux-preview-r22`, with a matching ZIP and
ZIP SHA256 sidecar. Its manifest records 42 empty user-data tables, all 46 source
standard modules matched to installed exports, native unload guards and 13
passing verification gates. Both schema-20 and schema-21 fixture transfers
passed exact workspace row/field comparison and reopen, with source/template
hashes unchanged. Unknown tables, populated templates and existing destinations
were rejected. The ZIP inventory and every packaged-file hash were verified.

The packager refuses existing destinations and publishes no GitHub release.
The distributed 0.21.0 checkpoint remains unchanged. Native manual acceptance
is required before this candidate becomes a replacement release.

## Remaining before redesign rollout

- Continue native visual review against the user's report-form density goals.
- Verify real laptop/desktop layouts and Windows 100%, 125% and 150% scaling;
  inspect responsive navigation, wrapping and remaining destination layouts.
- Manually verify keyboard navigation, focus and screen-reader behavior in
  Access; automated geometry checks do not establish accessibility.
- Focused review and Reports return are implemented in r23; complete manual
  review of the new screen and native return behavior.
- Complete current-template caption/block review, printed layout work and
  remaining interoperability/field mappings within their documented scope.
- Complete manual native acceptance before publishing a replacement release.
  Schema-20 transfer and exact-artifact checks are available in the separate
  local candidate package, with original workspaces preserved.

Do not transfer operational data into a seeded demo. Use a separate blank
build and preserve existing workspaces during migration.



Final r23 blank SHA256: `6C93E7F2F1F6D95A70FF62AC780770012FBCEBB06E88466434D7CAEAFEE59CB8`.
Final r23 demo SHA256: `46619E04F5A00F0B79A02569A15DA9465B4605944CFDA2741DB09D2CDD1992DA`.
