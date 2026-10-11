## 2026-10-10 - standardized comments input r104

Measured the supplied printed EVAL example and implemented fixed report-shaped comments input, locked plain-text Courier New and margins, narrow-view disclosure, and active-text counts. Overflow preserved. 129 native geometry/save assertions passed. Synthetic demo contains Lorem Ipsum; separate clean blank prepared without changing populated local r103. Exact printed fit remains pending. See NARRATIVE_LAYOUT_R104_2026-10-10.md.

## 2026-10-10 - r103 blank deployment prepared

Built latest source into a separate blank r103 file, compiled, verified empty user tables and setup startup, then passed exact-file scratch checks for Members, pending saves, approval retry and change-driven timers. README, deployment notes, verification and manifest updated for GitHub publication. Existing deployments/workspaces preserved.

## 2026-10-10 - VBA function comments added

Added purpose and decision comments to 117 functions/subroutines across nine source modules covering navigation, live validation, recovery, comparison queues, snapshots, defaults and transactional saves. Executable text is unchanged. Imported all nine modules into a scratch r102 blank database and compiled successfully. Deployment databases remain unchanged.

## 2026-10-10 - larger synthetic groups measured

Created a separate 525-report scale database: 25/50/100-report groups for each report type. Measured native group opens, row navigation, queued saves and editor opens on scratch copies. Row moves 5-7 ms; saves 93-110 ms; group opening 0.64-1.09 s; report editor opening 2.44-4.01 s. Exactly 90 tested edits produced 90 revisions and zero group-history rows. Next priority: profile editor startup before optimizing. See [scale results](SCALE_R102_2026-10-10.md). No production source or schema changes.

## 2026-10-10 - r102 efficiency priorities 3-6 complete

Completed targeted editor/group refresh caching; removed duplicate hidden comparison list and adapted proposal selection to the grid; routed Members directly and simplified group tab transitions; made validation/recovery timer work event-driven. Idle calls produce zero validation scans. Saved averages refresh after commit; typing and field-exit checks/recovery remain. All 49 user tables preserved; schema 25. See [implementation and verification](EFFICIENCY_R102_2026-10-10.md). Physical display/accessibility and scale performance verification remain outstanding.

## 2026-10-10 - r101 live Members and cached display guards

Completed single-user efficiency priorities 1 and 2. Members reads current reports directly and no longer creates historical snapshots on opening or direct cell saves. Row focus uses loaded state; authoritative checks run at save. Historical proposal baselines/conflict checks and pending-save recovery remain. Three group opens add zero history rows. All 49 user tables preserved; schema 25. See [implementation and verification](LIVE_MEMBERS_R101_2026-10-10.md). Next: target editor refreshes, prove/remove legacy comparison controls, then simplify group page navigation.

## 2026-10-10 - r100 performance candidate

Completed unchanged comparison baseline reuse, input-driven live validation caching, unchanged checkpoint write suppression, joined saved-average reads, and conditional comparison ITA writes. Three unchanged group opens add 8 rather than 24 history rows; timing speedup is not established. Pending-save safeguards retained; schema 25; all 49 user tables preserved. See [performance notes](PERFORMANCE_R100_2026-10-10.md). Successful cell-save snapshot creation remains a follow-up.

## 10 October — r99 review priority 1 completed

Queued-cell failure recovery now retains failed and later edits, offers Retry/Discard, cancels failed native closure, guards contributor attribution and prevents proposal/group operations bypassing pending cells. 42 failure-path checks pass, alongside existing save/approval/event/startup/UI suites and 49-table preservation. Details in PENDING_CELL_RECOVERY_R99_2026-10-10.md. Local r99 blank/demo/preserved workspace created; no publication. Priorities 2/3 (snapshot lifecycle and change-driven editor work) remain outstanding.

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

# Manual feedback implementation progress - 6 October 2026

Plan: MANUAL_FEEDBACK_CHANGE_PLAN_2026-10-06.md. Completed means implemented and functionally checked in a development candidate; physical display acceptance and publication remain separate.

## Current status - r75

- Implemented and functionally checked: items 1-15, 18-21 and the selected-group import implementation in 22. Native NAVFIT UI compatibility acceptance is still pending.
- Item 17: manual directory, stable identities, automatic links and explicit correction implemented; RUAD refresh behavior depends on item 16's sample mapping.
- Item 16: RUAD import awaits the synthetic sample path and its verified rows/columns.
- Item 23: measured native designer/runtime/print branches documented; font-locked equivalent wrapping and actual-output proof remain incomplete.
- Item 24: expanded to 72 rate/source collections and 20 excerpts; rate filters in both catalog and saved snippets, plus manual snippet rate assignment. Full collection text is linked rather than redistributed.
- Latest populated demo: output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r75.accdb, schema 23, with synthetic prior history. Published release unchanged. Physical display/scaling and screen-reader acceptance remain outstanding.

## Incremental log
| Item | State | Evidence |
|---|---|---|
| 1 Overview contributor dropdown | Complete | Native switch updates active profile and header; existing editor/comment attribution guards retained. |
| 2 Hide Access Objects at startup | Complete | Database startup preference and native hide action verified; no global Access setting changed. |
| 3 Snippet return to originating report | Complete | Report stays open, pending narrative survives, no revision created; Resources origin returns to Resources. |
| 4 Escape failed save with undo | Complete | Invalid SSN blocks save/leave; Undo restores committed narrative, clears session checkpoint and permits navigation. Save exceptions become inline feedback. |
| 8 Editable career recommendations | Complete | Both fields stage, validate, reload and apply through existing atomic comparison services; stale fields visibly locked. |

Phase 1: 20 native feedback assertions passed on build/ux-r33-fixed-blank.accdb. No new release published. Membership, defaults precedence, status propagation, prior-eval selection and naming semantics remain awaiting clarification; RUAD mapping awaits dummy sample.

Items 8, 9 and 11 complete: 19 career/table assertions verify both career fields stage, survive reload, reject oversized values and apply atomically; Reports shows ITA/promotion/groups, group members show ITA/promotion, and group-list defaults opens the section editor. Regression: 90 block/comparison and 41 group lifecycle assertions pass. Item 7 (default editable member grid and per-row report entry/return) is in progress.


Item 7 complete for populated groups: 12 native member-grid assertions verify default entry, per-row Open, saved return, selected member preservation and stale-career locks. Empty groups show tools and an explicit add-members message; archived groups retain inspection view.

Current candidate: output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r36.accdb; blank: build/ux-r36-verified-blank.accdb. Installed-source audit passes all 59 modules and native review/return event bindings. Broader layout/workflow checks are running. This is a development build, not a published release.


## Completed after r36

Working assumptions (pending user correction): Move replaces previous membership; global defaults supply editable starting values, group overrides global and report overrides group; group status does not propagate to reports. Export rank is report rank/rate or group paygrade, date is ISO format; mixed group ending dates use MIXED. Existing folder data and legacy defaults remain preserved.

- Items 5/6/12 complete: group chooser with Preview/Apply and stale guards, summary-group/Ungrouped filters, and journaled Initialized / In-Work / Ready for Signature / Completed statuses. 22 native assertions pass.
- Item 13 complete: global shared-block defaults, inherited values and Reset to inherited; existing saved reports are not silently rewritten. 26 native assertions pass; 90 block-tool regression assertions pass.
- Item 21 complete: directory chooser with automatic rank/end-date filenames and safe paired-baseline collision suffixes. 9 native assertions pass. Actual NAVFIT export regression remains to be rerun on the final candidate.
- Item 24 starter implementation complete: eight short source-linked NavyWriter examples, search, editable preview, attribution and duplicate detection. 11 native assertions pass. This is a starter catalog, not a copy of the full website.

r36 verification: 59 installed modules matched source; 36 demo, 1846 responsive-layout and 49 workflow assertions passed. Forced Access termination checks passed for valid save and invalid raw-value recovery. These results do not establish physical-display or screen-reader acceptance.

Latest incremental blank candidate: build/ux-r41-blank.accdb (64 modules). No release published. Item 14 report-header review actions is in progress. RUAD mapping still awaits the promised synthetic sample; directory/history/schema, notes, selected-group import and measured narrative fit remain outstanding.

Item 14 complete: report-header Review setup, Accept step, Return report and Local approval/reopen actions retain the report onscreen and use existing ordered-route services. Decisions request a note, operate on saved revisions, and preserve the existing pending-edits chooser. Hidden adapter is closed after success or rejection. 18 targeted assertions passed, including stale route rejection and compact header bounds. Availability failures show the route-specific reason inline; visible review windows are not silently closed with pending notes.

Current incremental blank candidate: build/ux-r42-blank.accdb. Full-field import investigation confirms that report mapping/provenance exists, but native folder-to-summary-group selection and atomic group creation still require implementation and tests (item 22 remains outstanding).

Item 14 final check: 22 assertions pass, including visible disabled states and stale-route explanations. Item 22 wording correction: 8 assertions distinguish full imports from legacy narrative imports and ignore undone receipts.

Item 22 group-selection implementation: 19 targeted assertions pass on an instrumented scratch build. Choose NAVFIT groups -> select type/paygrade partitions -> preview full fields and folder defaults -> atomically apply groups/reports -> results. Existing drafts are not overwritten; duplicate receipts require returned-file reconciliation. Failed late partition rolls back all selected groups, reports, receipts and defaults. A DAO defaults lookup was corrected to see newly created groups inside the outer transaction. Final installed-build verification and native round-trip proof remain pending; item 22 is not yet marked fully verified.

r44 assembled checks: 65 source modules, 36 demo, 1972 responsive layout, 27 feedback, 19 career/table, 12 member-grid, 22 organization, 26 global defaults, 9 filename and 11 NavyWriter assertions passed. r45 adds 8 import-scope assertions; 49 workflow checks and forced shutdown/recovery passed; 90 block-tools, 41 group lifecycle, 39 group correctness and 59 live validation checks passed. NAVFIT adapter 497 and group export 867 checks also passed before the group-selection addition.

Item 23 inspection in progress: 229 installed methods inspected without invoking target code. All three designer comments controls initialize Courier New 12pt; print pitch branches change font handling, so designer dimensions alone are not printed-fit proof. Evidence and limitations recorded in NAVFIT_PRINT_ENGINE_FINDINGS.md. No new truncation/font restriction imposed from partial evidence.

Item 22 implementation complete in r49, with compatibility acceptance still separate: selected native-group import and full mapped import/export data round trips passed 244 checks on EVAL/CHIEFEVAL/FITREP. Includes late report-validation rollback, selection-change/duplicate guards, group defaults and provenance. Mismatched type/paygrade is rejected during source inspection. Native UI validation and physical printing remain unverified.

Current candidates: output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r49.accdb and build/ux-r49-blank.accdb, schema 21. Published release unchanged. Remaining plan items: 10, 15, 16, 17, 18, 19, 20 and 23, plus physical/native compatibility acceptance. Item 24 is a starter catalog only. RUAD sample path requested; mapping cannot be verified without it.

Item 15 service/UI implementation complete; final assembled-build verification in progress: posted contributor/note entries are append-only, with a full-text viewer, separate raw-text checkpoints, explicit Post / Close and keep checkpoint / Discard actions and owned-checkpoint recovery. Notes do not change report revisions or NAVPERS content. Legacy attribution is explicitly unknown. Checks: 19 service and 18 UI assertions pass. Schema-21 -> 22 demo transfer preserves all original data and creates nine legacy entries; schema-22 -> 22 transfer preserves two entries and one unposted checkpoint exactly. Item 20 notes portion is complete; directory/history schema still awaits its own design.

Item 15 final targeted checks: 19 service and 18 UI assertions pass, including Overview -> unposted notes -> reopen report/composer -> post. Access shutdown retains raw unposted note text without posting it. A schema-22 -> 22 transfer preserves both posted notes and the unposted checkpoint exactly. Current development candidates: output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r56.accdb and build/ux-r56-notes-blank.accdb, schema 22. Published schema-21 release is unchanged. See SCHEMA22_REPORT_NOTES.md for justification and data-preservation rules. Broad assembled-build checks are running.

r56 assembled verification passed: all 68 source modules match the installed demo; 1972 layout/focus/identity checks; 49 native-close/bulk/review checks; forced report shutdown/recovery; raw unposted note shutdown; 19 note services and 18 note UI/recovery checks. Schema-21 -> 22 and schema-22 -> 22 transfers preserve original data and notes/checkpoints. Item 15 is complete. Item 20 is complete for notes only; directory/history schema is next. Manual directory/history implementation is starting independently of the pending RUAD sample.

### Command directory and prior-history screens (schema 23)
- Implemented manual member profiles in Resources, automatic directory linking on report saves, written-report history, and manual prior EVAL/CHIEFEVAL/FITREP capture (period, promotion, seven traits, source context).
- Verified: 17 service identity/history assertions and 13 native Access UI assertions. Invalid identity input stays in the profile editor; names remain separately stored and are never truncated.
- Schema 23 adds four directory/history tables; report fields and existing revisions are unchanged. Migration verification is in progress. Latest tested UI foundation: r61 blank.
- Prior matching uses the latest evaluation ending strictly before current report start. Equal latest ending dates show Ambiguous; unconfirmed identity shows Unavailable. RUAD import still awaits the synthetic sample path.
- Summary-group prior ITA/promotion columns are under implementation; bulk creation from directory is pending.

### r66 directory/history completion
- Items 17, 18 and 19 implemented: manual command directory, automatic report identity links, prior evaluation capture, and selected-member report creation using inherited global/group defaults.
- Item 10 implemented: read-only prior ITA and promotion columns in the editable summary-group member grid, with period/source context. Unconfirmed identity and ambiguous history are explicit.
- Item 20 directory/history schema portion implemented: additive schema 23; schema 22-to-23 and 23-to-23 transfers preserve original rows/fields. Existing published schema-21 release is unchanged.
- Verification: 20 atomic batch assertions and 16 native directory/history UI assertions passed, including late-failure rollback, stale inputs, all three report types and explicit identity confirmation. Earlier directory/history services: 17 assertions; prior member grid: 14 assertions.
- Populated candidate: output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r66.accdb. Nine synthetic reports, three groups and nine directory members. Reopened schema-23 transfer verified all original rows/fields and unchanged source hash.
- Remaining implementation: RUAD file mapping/import awaits the synthetic sample path; exact narrative print-fit behavior remains under investigation. Physical display, screen-reader and native NAVFIT round-trip acceptance remain separate outstanding checks.

r66 additional checks: 21 directory/history assertions now cover confirmed and provisional identities, NOB/incomplete ITA, tied latest dates, archived external history and strict prior-period selection. The 49 workflow checks pass on the blank fixture. Narrative inspection now covers 237 methods and identifies separate runtime editing fonts and all three comments print branches; item 23 remains incomplete because native print equivalence has not been established.

### r67 clean-build and populated-history candidate
- A clean build recreates 63 forms and 1,950 controls. All 72 installed source modules match; native review/return bindings and VBA compile verified. Its final hash calculation initially encountered an Access file lock after database close; moved hashing after Access shutdown. The generated artifact was subsequently reopened and independently audited successfully.
- Native clean-build checks: 20 directory batch, 16 directory/history UI and 244 selected-group import assertions pass. r66 populated responsive checks: 1,972 pass.
- New populated demo: output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r67.accdb. Schema-23 transfer preserves every original row/field and source hash. Added eight explicitly synthetic prior evaluations without changing current report revisions. The remaining provisional identity deliberately has no inferred history. Repeated demo seeding adds zero duplicates.
- The populated-demo suite passes all 36 behavior assertions. Its first source-hash check overlapped the deliberate seeding operation; rerun sequentially to verify the deliverable remains unchanged.
- No publication or Git push performed. Remaining implementation: verified RUAD mapping/import and exact narrative-fit behavior. Native display/screen-reader and NAVFIT UI compatibility acceptance are outstanding.

### r68 identity correction
- Identity confirmation now lets users choose the correct active directory member, previews their name/rank and requires a reason. Selection changes and stale member versions cannot apply an old preview. Saved report values and revisions remain unchanged.
- 21 native directory/history UI assertions pass, including member reassignment, stale preview rejection and unchanged report revision; all 72 installed source modules match.
- Latest demo: output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r68.accdb; blank: build/ux-r68-identity-blank.accdb. Schema-23 transfer preserves all existing rows/fields, including eight prior records and 56 trait rows; source hash unchanged.
- r67 clean-build follow-up: 21 history services, 14 prior-grid and 49 workflow assertions pass; forced Access shutdown saves valid edits and retains invalid raw edits for recovery.

r68 final artifact check: all 72 installed source modules match, native event bindings and compilation pass, and 36 populated-demo assertions pass with the deliverable hash unchanged. Current test processes closed; existing user Access instances were retained.

## Narrative and policy increment â€” r69â€“r75

Rate catalog/filter completed; 29 native assertions and 78 distinct source URLs checked successfully. Expanded live BUPERS checks, official counseling status save/PDF/native exchange, and source-page verification implemented. 37 verified citations; 102 live checks, 170 service checks, 12 new counseling-PDF checks and existing exchange/PDF suites pass. Policy scope, remaining gaps and eight researched leadership-tool recommendations: NARRATIVE_POLICY_REVIEW_2026-10-06.md. Final artifact audit and demo acceptance recorded in CURRENT_STATE.md. Schema remains 23; publication and physical/native NAVFIT UI acceptance remain separate.

Final r75 acceptance: 72 installed modules match current source; 102 live-validation, 36 populated-demo, 30 guide and 29 populated catalog assertions pass. The schema-23 transfer preserves every original row/field and source fingerprint. No Git push or release publication.

## Deployment 0.24.0-preview

Release regression checks exposed and fixed the backup worker’s schema-21 ceiling and approval stamps that became stale solely from approval edit locks. Review currency now excludes member lock changes while bulk/edit stamps retain them. Major workflow tests include sibling approvals; 46 major and 39 group-correctness assertions pass before final exact-package verification. Migration tools and release guides target schema 23; exact release evidence is in DEPLOYMENT_0.24.0_VERIFICATION.json.

## Additional feedback completion — r86

The subsequent admin/comparison feedback is implemented in 0.25.0-dev / schema 24. Combined Admin data includes Block 28; redundant canonical-name/readiness subfields are hidden without losing background values. Direct Defaults/Comparison, member double-click opening, immediate comparison saves with approval reopening, full-height trait guidance, My Work contributor and populated rate filters are implemented. Reviewed-proposal application uses owned DAO database references and passes the native regression.

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r86.accdb`. This supersedes r75 for development testing; published 0.24.0 remains unchanged. Focused checks: 142; responsive: 2,113; exact-demo journeys: 36. Detailed evidence and remaining acceptance items: `ADMIN_COMPARISON_FEEDBACK_2026-10-06.md`.
