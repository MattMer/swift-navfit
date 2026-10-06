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
