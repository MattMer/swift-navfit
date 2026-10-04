# Current delivery - three structured name features

Swift NAVFIT 0.20.0-prototype, schema 19, output/Swift_NAVFIT_0.20.0-structured-names.accdb. Adds member first/middle/last/suffix, reporting-senior last/initials and counselor last/initials. All ten parts are explicit and separate from existing full names. Apply checkpoints pending values; Save draft commits all 78 fields atomically. See STRUCTURED_NAMES_GUIDE.md.

Verification: 281 new assertions; 2001 total assertions, 218 reference checks, nine layout checks and all 2472 retained import values. Schema 18-to-19, 19-to-19 and 5-to-19 migrations preserve original rows/fields, source hashes and reopen. All 20 distribution user-data tables are empty; delivery matches the tested candidate. An initial build exceeded Access's form-height limit; the final editor uses one compact details-button row. The rollback fixture was corrected to attempt a changed save. Runtime/x86, comprehensive visual inspection, Navy rules and full interchange remain pending.

# Current delivery - senior identifiers, contact and raters

Swift NAVFIT 0.19.0-prototype, schema 18, output/Swift_NAVFIT_0.19.0-report-people.accdb. Adds reporting-senior SSN/DoD ID, structured address/contact fields, and rater/senior-rater names/dates. Apply checkpoints pending values; Save draft commits all 68 fields atomically. Leading zeros and incomplete entries are preserved. Names/dates do not establish signatures or approval. See REPORT_PEOPLE_GUIDE.md.

Verification: 364 new behavioral assertions; 1720 total assertions, 218 reference checks, six layout checks and all 2472 retained import values. Schema 17-to-18, 18-to-18 and 5-to-18 migration checks preserve original rows/fields, source hashes and reopen. The distribution's 20 user-data tables are empty and its binary matches the tested candidate. Runtime/x86, comprehensive visual inspection, policy rules and full interchange remain pending.

# Current delivery - recommendations, retention and statement

Swift NAVFIT 0.18.0-prototype, schema 17, output/Swift_NAVFIT_0.18.0-decision-details.accdb. Adds assignment recommendations A/B and four explicit Yes/No/blank source-box selections for retention and member statement. Apply checkpoints pending values; Save draft commits all 55 fields atomically. No signatures, approval or policy inference. See DECISION_DETAILS_GUIDE.md.

Verification: 198 new behavioral assertions; 1356 total behavioral assertions, 218 reference checks and all 2472 retained import values. Schema 16-to-17, 17-to-17 and 5-to-17 migration checks preserve original rows/fields, source hashes and reopen. The editor button was moved to a separate row to prevent period-field overlap; focused layout bounds and overlap checks verify the final build. Runtime/x86, comprehensive visual inspection and full interchange remain pending.

# Current delivery - counseling, readiness and billet

Swift NAVFIT 0.17.0-prototype, schema 16, output/Swift_NAVFIT_0.17.0-service-details.accdb. Adds counseling date/name, readiness summary/detail/date and billet subcategory. Apply checkpoints pending values; Save draft commits all 49 fields atomically. Editable reference suggestions do not enforce Navy policy. See SERVICE_DETAILS_GUIDE.md.

Verification: 1158 behavioral assertions, 218 reference checks and all 2472 retained import values passed. Schema 15-to-16, 16-to-16 and 5-to-16 migrations preserve original rows/fields, source hashes and reopen. All 20 distribution user-data tables are empty. Runtime/x86, comprehensive visual inspection and full interchange remain pending.

# Current delivery - three member-detail features

Swift NAVFIT 0.16.0-prototype, schema 15, output/Swift_NAVFIT_0.16.0-member-details.accdb. Adds SSN/DoD ID text entry, four Yes/No/blank component selections and canonical date reported in one details panel. Apply checkpoints pending values; Cancel preserves the editor. Save draft commits all 43 fields atomically. Incomplete identifiers/blanks remain saveable; official identity/component/date policy and external mapping remain pending. See MEMBER_DETAILS_GUIDE.md.

Verification: 125 new assertions plus 881 prior assertions, 169 catalogue checks and 2472 retained import values passed. Schema 14-to-15, 15-to-15 and 5-to-15 migrations preserve original rows/fields, source hashes and reopen. All 20 distribution user-data tables are empty; delivery matches tested candidate. Full NAVFIT mapping/export, Navy validation, Runtime/x86 and comprehensive visual verification remain pending.

# Current delivery - three report-context features

Swift NAVFIT 0.15.0-prototype, schema 14, output/Swift_NAVFIT_0.15.0-report-context.accdb. Adds report occasions (four Yes/No/blank flags), report categories (three flags), and primary duty/multiline duty description in separate panels. Apply checkpoints only that panel; Cancel preserves the editor. Save draft commits all 36 fields atomically. No inference or combination/eligibility validation. See REPORT_CONTEXT_GUIDE.md.

Verification: 195 new assertions plus 686 prior assertions, 169 catalogue checks and all 2472 retained import values passed. Schema 13-to-14, 14-to-14 and 5-to-14 migrations preserve original rows/fields, source hashes and reopen. All 20 distribution user-data tables are empty; delivery matches tested candidate. An initial harness compile error (Long used in For Each) was corrected before successful testing. No failed test candidate is published. Full NAVFIT mapping/export, Navy validation, Runtime/x86 and comprehensive visual verification remain pending.

# Current delivery - reporting senior details

Swift NAVFIT 0.14.0-prototype, schema 13, output/Swift_NAVFIT_0.14.0-reporting-senior.accdb. Adds name/grade/designator/title/UIC in a separate panel. Apply transfers pending values and checkpoints them; Cancel preserves the editor. Save draft commits all 27 fields atomically. Older unrecorded values restore to null; new snapshot rows are required. See REPORTING_SENIOR_GUIDE.md.

Verification: 81 new assertions and all 605 prior assertions passed, plus 169 catalogue checks and 2472 retained import values. Schema 12-to-13, 13-to-13 and 5-to-13 migrations preserve original rows/fields, source hashes and reopen. All 20 distribution user-data tables are empty; delivery matches tested candidate. Reporting authority/eligibility and external mapping/export remain pending, as do Runtime/x86 and comprehensive visual verification.

# Current delivery - promotion status

Swift NAVFIT 0.13.0-prototype, schema 12, output/Swift_NAVFIT_0.13.0-promotion-status.accdb. Adds REGULAR/FROCKED/SELECTED/SPOT selection, distinct from promotion recommendation. All 22 draft fields participate in snapshots, comparison, restore, duplication and recovery; status edits stale prior reviews. Legacy unrecorded status restores to null. Incomplete drafts remain saveable. See PROMOTION_STATUS_GUIDE.md.

Verification: 66 new assertions plus 539 prior assertions, 169 catalogue checks and all 2472 retained import values passed. Schema 11-to-12, 12-to-12 and 5-to-12 migrations preserve original rows/fields, source hashes and reopen. All 20 distribution user-data tables are empty; delivery matches tested candidate. Eligibility/policy and external mapping/export remain pending, as do Runtime/x86 and comprehensive visual verification.

# Current delivery - explicit report assessment selections

Swift NAVFIT 0.12.0-prototype, schema 11, output/Swift_NAVFIT_0.12.0-assessment.accdb. Adds manually selected whole-report Observed/NOB and the six NAVFIT reference promotion recommendation labels. Blank remains saveable. All 21 draft fields participate in snapshots, comparison, restore, duplication and recovery; assessment edits stale earlier reviews. No inference from traits and no policy/eligibility validation. See ASSESSMENT_GUIDE.md.

Verification: 70 new assertions and all 469 prior assertions passed, plus 169 catalogue checks and all 2472 retained import values. Schema 10-to-11, 11-to-11 and 5-to-11 migrations preserve original rows/fields, source hashes and reopen. Legacy reports/checkpoints have null new fields. All 20 distribution user-data tables are empty; delivery matches tested candidate. Runtime/x86 and comprehensive visual verification remain pending.

# Current delivery - trait summaries

Swift NAVFIT 0.11.0-prototype, schema 10, output/Swift_NAVFIT_0.11.0-trait-summary.accdb. Read-only editor and latest-saved-revision arithmetic summaries show numeric average/count, NOB count and blank count. Blanks label an average partial; no numeric marks or invalid input yield unavailable. Save and checkpoint recovery refresh the summaries. No schema or derived-value storage change. See TRAIT_SUMMARY_GUIDE.md.

Verification: 44 new assertions and all 425 existing assertions passed, plus 169 catalogue checks and all 2472 retained import values. Populated schema-10 migration preserved original rows/fields, source hash and reopen. All 20 user-data tables in the distribution are empty; delivery matches the tested candidate. Runtime/x86 and comprehensive visual checks remain pending. Observation/recommendation rules, group context and official validation remain pending.

# Current delivery - form-specific trait editor

Swift NAVFIT 0.10.0-prototype, schema 10, output/Swift_NAVFIT_0.10.0-traits.accdb. Seven traits per form have reference descriptions, distinct blank/NOB values and marks 1.0–5.0. All 19 draft fields participate in snapshots, comparison, restore, duplicate and recovery. Trait edits stale earlier review assignments. See TRAITS_GUIDE.md.

Verification: 425 behavioral assertions, 169 catalogue checks, three narrative imports and all 2472 retained-value checks passed. Schema 9-to-10, 10-to-10 and 5-to-10 migrations preserve original data, source hashes and reopen. All 20 user-data tables are empty; distribution matches tested candidate. Runtime/x86 and comprehensive visual checks remain pending. Averages, recommendations, official validation and external trait mapping/export remain pending.

# Current delivery - administrative draft fields

Swift NAVFIT 0.9.0-prototype, schema 9, output/Swift_NAVFIT_0.9.0-administration.accdb. Adds manually entered official full name, grade/rate, designator/designation, unit UIC and ship/station with measured storage limits and no truncation. All 12 report fields participate in snapshots, comparison, restore/duplicate and recovery. Administrative edits stale older review assignments. Missing legacy fields display Not recorded; restore clears them to null. New snapshots have a schema marker and reject missing administrative rows. See ADMINISTRATION_GUIDE.md.

Verification: 53 administrative assertions and all existing focused suites pass (350 assertions total), plus three narrative imports and exact 2472 retained-value checks. Schema-8-to-9, schema-9-to-9 and schema-5-to-9 migrations preserve every original field, including old checkpoints, and reopen; source hashes unchanged. All 20 user-data tables in the master are empty; delivery byte-identical to tested candidate. Full visual/Runtime/x86, administrative import mapping and complete NAVFIT export remain pending. The supplied NAVFIT bundle was inspected read-only; lookup/template findings are documented in NAVFIT_PACKAGE_FINDINGS.md.

# Current delivery - review assignments and My Work

Swift NAVFIT 0.8.0-prototype, schema 8, output/Swift_NAVFIT_0.8.0-review-routing.accdb. Adds one pending local review per report, revision references, requester/reviewer attribution, due dates, My Work, contributor creation/switching and accept/return/cancel event history. Later revisions block outcomes on stale assignments. Draft editing remains available; internal acceptance is not official approval. Sequential route templates and official locks remain pending. See REVIEW_ROUTING_GUIDE.md.

Verification: 38 routing, 54 workflow, 28 comments, 24 provenance and 46 smoke assertions pass. Late event-write failures roll back task creation/outcome changes. Schema-7-to-8 and schema-8-to-8 migrations compare all original rows/fields and reopen; source hashes unchanged. All 20 user-data tables in the master are empty; delivery byte-identical to tested candidate. Cleanup no longer closes already-closed recordsets; rollback precedes cleanup. Full visual/Runtime/x86 verification remains pending; no export or policy-validation capability added.

# Current delivery - report review comments

Swift NAVFIT 0.7.0-prototype, schema 7, output/Swift_NAVFIT_0.7.0-review-comments.accdb. Adds saved-revision threads, replies and resolve/reopen events with local contributor attribution. Posted comments have no edit/delete actions, remain separate from official narratives and never create report revisions. Archived reports retain readable history and reject additional activity. Return preserves pending editor values. Routing, field anchors and approval tracking remain pending. See REVIEW_COMMENTS_GUIDE.md.

Verification: 28 comment assertions, 54 workflow assertions, 24 provenance assertions and 46 smoke checks pass. Schema-6-to-7 and schema-7-to-7 migrations compare all original rows/fields and reopen; source hashes unchanged. All 18 user-data tables in the master are empty; delivery byte-identical to tested candidate. Full visual/Runtime/x86 verification remains pending; no export or policy-validation capability added.

# Current delivery - import provenance

Swift NAVFIT 0.6.1-prototype, schema 6, output/Swift_NAVFIT_0.6.1-provenance.accdb. Read-only receipt/source-value viewer with table, zero-based row and field-name filters, complete memo display, null/empty markers and receipt metadata. Uses retained data without reopening external source files. See PROVENANCE_GUIDE.md.

Verification: 24 provenance assertions, 54 integrated workflow assertions and 46 smoke checks pass. Schema-6-to-6 migration compares all original rows/fields and reopens. All 16 user-data tables in the master are empty; delivery byte-identical to tested candidate. Full visual/Runtime/x86 verification remains pending. No additional official-field mapping or export implementation is claimed.

# Current delivery — workflow tools

`output/Swift_NAVFIT_0.6.0-workflow-tools.accdb`, app 0.6.0-prototype, schema 6. The remaining recommended workflow additions now have initial implementations: saved draft completeness checklist, explicit two-way narrative reconciliation, attributed searchable snippets, three contributor-persisted library column presets, bulk folder preview/apply, guarded import/batch undo and operation journal. Recovery and revision comparison continue to work. See WORKFLOW_TOOLS_GUIDE.md for steps and precise limits.

Verification: 54 integrated assertions, all existing draft/narrative/comparison/recovery/direct-review suites, and exact import/provenance tests pass. Schema-5-to-6 and schema-6-to-6 migrations preserve every original row/field and source hashes; targets reopen. All 16 user-data tables in the delivery are empty. The published master is byte-identical to the tested candidate; MANIFEST.json records its hash.

Testing repaired a DAO parameter-object use in snippet search, stale domain-lookup reads of applied revision IDs inside nested transactions, and an undo recordset opened before its transaction. Native recordset reads and appropriate transaction boundaries now pass exact undo and injected bulk-failure checks. No failed candidate is the current delivery.

Scope remains limited: narrative-only mapping/reconciliation, local draft checks rather than Navy validation, predefined column views rather than arbitrary custom columns, and folder moves rather than official-field bulk edits. Export, full official field models/grouping, routing/comments, structured PDF import and printed-layout verification remain pending. Full visual inspection and Runtime/x86 are untested.

## Prior milestones
# Current milestone — pending-edit recovery

Current delivery: `output/Swift_NAVFIT_0.5.0-recovery-tested.accdb`, app `0.5.0-prototype`, schema 5. The draft editor checkpoints changed pending values every 30 seconds while Access can process timer events. A manual checkpoint button sits below the narrative section. It captures active textbox text before focus leaves, including incomplete date entries, and retains long narratives. Unchanged editors do not generate checkpoints.

Home → Recover pending edits lists the active contributor's checkpoints. Recover reopens pending edits in the editor; it never commits them automatically. Save draft creates a normal revision and clears that session's checkpoint atomically. Explicit editor close/discard removes its checkpoint; closing Access without the explicit discard action retains the last checkpoint. Discard checkpoint in the recovery screen requires confirmation.

If the underlying report has a newer revision, normal recovery is rejected and Recover as new draft creates a separate editable draft on subsequent save. Archived/locked reports also require recovery as new drafts. The editor now checks its baseline revision before saving to prevent overwriting newer committed changes. No single-user lock or strong profile authentication is claimed.

Validation: 34 recovery assertions across two Access sessions, including new drafts, long text, invalid dates, rollback retaining a prior checkpoint, save/discard cleanup and stale revision protection. An accelerated timer test verifies the actual OnTimer callback captures active textbox text; the production interval is checked as 30000 ms. Also 46 regression checks, 39 narrative assertions, 22 comparison assertions and full import/provenance tests pass. Schema-4-to-5 and schema-5-to-5 migrations preserve every original field, including two recovery checkpoints; source hashes unchanged; targets reopen. All 12 user-data tables in the blank master are empty.

Limits: these are pending-edit checkpoints, not database backups. Edits since the last successful checkpoint can be lost; modal dialogs or a busy Access session can delay timers. Recovery requires a readable workspace file. Tests simulate interrupted editing by retaining checkpoints, then closing/reopening in a new Access process; abrupt process termination, database corruption and power-loss durability were not tested. Full visual inspection, Runtime/x86 and older-schema-to-5 migration remain pending. Printed layout remains deferred.

## Prior milestones
# Current milestone — revision comparison

Current delivery: `output/Swift_NAVFIT_0.4.1-revision-compare-tested.accdb`, app `0.4.1-prototype`, schema 4. In the draft editor select a saved revision, then **Compare with latest saved**. A read-only screen shows the selected and latest saved snapshots side by side. The changed-fields filter is on initially; turn it off to inspect all seven fields. Select a field to display its complete saved values. No automatic save or restore occurs; unsaved editor entries are excluded and preserved on return.

Comparison distinguishes null, empty text, case-only changes and narrative fields absent from pre-narrative snapshots. Missing required older fields or inconsistent snapshot null flags reject comparison. Values are displayed in locked textboxes, including long multiline narratives. This initial comparison is at field level; inline word/line highlighting and selection of two arbitrary revisions through the UI remain pending. The service supports explicit pairs of revisions from the same report; cross-report pairs are rejected.

Validation: 22 comparison assertions pass, including active form/return, unsaved preservation, no revision writes, long narrative fidelity, all/changed filters, same-revision behavior, cross-report rejection and legacy missing narrative markers. Also 46 regression checks, 39 narrative assertions, 12 direct-review checks and the full import/provenance/rollback suite pass against the exact artifact. All 11 user-data tables are empty in the delivery. Full visual inspection and Runtime/x86 remain pending. Schema is unchanged; migration into this build has not been repeated this milestone.

## Prior milestones
# Build status — 3 October 2026

Current delivery: `output/Swift_NAVFIT_0.4.0-narrative-import.accdb`, app `0.4.0-prototype`, schema 4. Deliberate partial NAVFIT import creates a new Swift draft with the recognized form kind and confirmed Comments narrative mapping. The review screen accepts an optional user-entered local member display name and requires a confirmation explaining scope. Dates, traits, folders and other official fields remain unmapped. The editor marks imported drafts accordingly.

Usage: complete workspace setup; Home → Review NAVFIT database → select report → inspect Comments → optionally enter a local member display name → Create draft from narrative → confirm. Open the draft library to continue. It never updates an existing report. Exact same source-file SHA256/report-ID imports are rejected; unknown/null types remain quarantined.

ImportReceipt retains file hash/path, source report ID, author, time, adapter version and value encoding. ImportRawValue retains every data-table field from the source, with table/row/field ordinal, field name, DAO type, null flag and value. Text is kept verbatim; dates are OLE numeric values and scalar values use VBA CStr with decimal mark recorded. These retained values are provenance, not an export adapter. Each receipt currently stores a full source snapshot, which duplicates data across selected-report imports; shared batch storage and a user-facing provenance/history viewer remain pending. Original and staged copies have different hashes and changed files may repeat reports, so cross-file duplicate matching/conflict reconciliation remains pending and is disclosed at confirmation.

Transactions include the new draft, its initial revision, receipt and provenance. Windows native SHA256 matches PowerShell; source is opened read-only and its hash checked before/after capture. A source change rejects the import. No optional backend, PowerShell runtime or external hashing tool is required by the application.

Validation: three typed imports, quarantine, exact-file duplicates, late injected provenance failure with verified rollback, reopen and all 2472 retained values across three imports pass. Also 46 regression checks, 39 narrative assertions and 12 direct-review checks pass. Schema-3-to-4 and schema-4-to-4 upgrades preserve every original row/field including 2472 provenance records; source hashes unchanged; migrated files reopen. All 11 user-data tables in the delivery are empty. Runtime/x86 and complete visual testing remain pending. UI confirmation wiring is checked; the modal confirmation was not automated end to end.

This is narrative-only import, not complete NAVFIT report import/export. Administrative models, field mappings, candidate duplicates, update/merge imports, recoverable import batches and export remain pending. Printing is deferred at user request.

## Previous delivery: 0.3.0
Current delivery: `output/Swift_NAVFIT_0.3.0-official-narratives.accdb`, app `0.3.0-prototype`, schema 3. Added an official narrative field to the editor: EVAL block 43, CHIEFEVAL block 40, FITREP block 41. Scroll below revision history to reach the narrative section, which has its own save button (saves the entire draft) and character count. General draft notes stay separate. Changing type before first save changes the narrative label; report type remains fixed after creation.

OfficialComments is Long Text and participates in atomic report/revision/change writes, duplication and restore. Native recordset history writes preserve long multiline content; binary comparisons detect case-only edits. Omitted narrative values in older save calls preserve the stored value, including null versus empty. Historical schema-1/2 snapshots have no official narrative; restoring them clears it to null rather than inventing content.

Validation: 39 narrative assertions across all three forms plus reopen/snapshot persistence, 46 regression checks and 12 direct NAVFIT review checks pass. Populated schema-2-to-3 and schema-3-to-3 migrations compare every original row/field, reopen successfully and preserve source hashes. Legacy draft notes are not automatically converted to official narratives. No full visual layout, Runtime/x86, narrative printed-fit or export compatibility claim is made. Schema-1-to-3 migration has not been tested this milestone.

This is the narrative portion of the official field model. Administrative blocks, traits, recommendations, deliberate import commit, policy validation, routing/comments and NAVFIT export remain pending. Printed layout is deferred at the user's request.

## Previous delivery: 0.2.3
Current delivery: `output/Swift_NAVFIT_0.4.0-narrative-import.accdb`, app `0.4.0-prototype`, schema 4. Deliberate partial NAVFIT import creates a new Swift draft with the recognized form kind and confirmed Comments narrative mapping. The review screen accepts an optional user-entered local member display name and requires a confirmation explaining scope. Dates, traits, folders and other official fields remain unmapped. The editor marks imported drafts accordingly.

Usage: complete workspace setup; Home → Review NAVFIT database → select report → inspect Comments → optionally enter a local member display name → Create draft from narrative → confirm. Open the draft library to continue. It never updates an existing report. Exact same source-file SHA256/report-ID imports are rejected; unknown/null types remain quarantined.

ImportReceipt retains file hash/path, source report ID, author, time, adapter version and value encoding. ImportRawValue retains every data-table field from the source, with table/row/field ordinal, field name, DAO type, null flag and value. Text is kept verbatim; dates are OLE numeric values and scalar values use VBA CStr with decimal mark recorded. These retained values are provenance, not an export adapter. Each receipt currently stores a full source snapshot, which duplicates data across selected-report imports; shared batch storage and a user-facing provenance/history viewer remain pending. Original and staged copies have different hashes and changed files may repeat reports, so cross-file duplicate matching/conflict reconciliation remains pending and is disclosed at confirmation.

Transactions include the new draft, its initial revision, receipt and provenance. Windows native SHA256 matches PowerShell; source is opened read-only and its hash checked before/after capture. A source change rejects the import. No optional backend, PowerShell runtime or external hashing tool is required by the application.

Validation: three typed imports, quarantine, exact-file duplicates, late injected provenance failure with verified rollback, reopen and all 2472 retained values across three imports pass. Also 46 regression checks, 39 narrative assertions and 12 direct-review checks pass. Schema-3-to-4 and schema-4-to-4 upgrades preserve every original row/field including 2472 provenance records; source hashes unchanged; migrated files reopen. All 11 user-data tables in the delivery are empty. Runtime/x86 and complete visual testing remain pending. UI confirmation wiring is checked; the modal confirmation was not automated end to end.

This is narrative-only import, not complete NAVFIT report import/export. Administrative models, field mappings, candidate duplicates, update/merge imports, recoverable import batches and export remain pending. Printing is deferred at user request.

## Previous delivery: 0.3.0
Current delivery: `output/Swift_NAVFIT_0.2.3-direct-review-tested.accdb`, app `0.2.3-prototype`, schema 2. Home now opens either an original NAVFIT database or a staged batch for read-only all-field review. The recognized shape is checked using ordered field names/types/sizes for Reports, Folders and Summary; linked/hidden tables are rejected. Report classification is derived directly from ReportType with binary comparisons instead of trusting a staging queue. No records are committed or exported.

Validation: 46 regression checks plus 12 direct-source and 12 staged-source review checks pass (70 assertions across suites). Original TEST_DB and staged-file hashes remain unchanged. Altered schemas are rejected. Full visual inspection, Runtime and migration into this build are not repeated/verified. Schema remains 2.

User confirmed TEST_DB is entirely synthetic, authorized changes, and reported Help/About 2.2.0.332 released August 9, 2026. Read-only executable metadata independently agrees on version 2.2.0.332. NAVFIT_BLANK.accdb is now supplied; its three data tables match the ordered field shape but include one folder and one report default row. Their meaning remains unproven.

A development mapping probe at `build/navfit-proof/NAVFIT_MAPPING_PROBE_TESTED.accdb` changes only Comments on the three typed reports. All 824 cells were reconciled, exactly three changes verified, and the original hash preserved. NAVFIT display, validation and printing are pending user inspection. This probe is copied from the synthetic fixture; it is not a Swift export or a blank production template.

## Previous delivery: 0.2.2
Current delivery: `output/Swift_NAVFIT_0.4.0-narrative-import.accdb`, app `0.4.0-prototype`, schema 4. Deliberate partial NAVFIT import creates a new Swift draft with the recognized form kind and confirmed Comments narrative mapping. The review screen accepts an optional user-entered local member display name and requires a confirmation explaining scope. Dates, traits, folders and other official fields remain unmapped. The editor marks imported drafts accordingly.

Usage: complete workspace setup; Home → Review NAVFIT database → select report → inspect Comments → optionally enter a local member display name → Create draft from narrative → confirm. Open the draft library to continue. It never updates an existing report. Exact same source-file SHA256/report-ID imports are rejected; unknown/null types remain quarantined.

ImportReceipt retains file hash/path, source report ID, author, time, adapter version and value encoding. ImportRawValue retains every data-table field from the source, with table/row/field ordinal, field name, DAO type, null flag and value. Text is kept verbatim; dates are OLE numeric values and scalar values use VBA CStr with decimal mark recorded. These retained values are provenance, not an export adapter. Each receipt currently stores a full source snapshot, which duplicates data across selected-report imports; shared batch storage and a user-facing provenance/history viewer remain pending. Original and staged copies have different hashes and changed files may repeat reports, so cross-file duplicate matching/conflict reconciliation remains pending and is disclosed at confirmation.

Transactions include the new draft, its initial revision, receipt and provenance. Windows native SHA256 matches PowerShell; source is opened read-only and its hash checked before/after capture. A source change rejects the import. No optional backend, PowerShell runtime or external hashing tool is required by the application.

Validation: three typed imports, quarantine, exact-file duplicates, late injected provenance failure with verified rollback, reopen and all 2472 retained values across three imports pass. Also 46 regression checks, 39 narrative assertions and 12 direct-review checks pass. Schema-3-to-4 and schema-4-to-4 upgrades preserve every original row/field including 2472 provenance records; source hashes unchanged; migrated files reopen. All 11 user-data tables in the delivery are empty. Runtime/x86 and complete visual testing remain pending. UI confirmation wiring is checked; the modal confirmation was not automated end to end.

This is narrative-only import, not complete NAVFIT report import/export. Administrative models, field mappings, candidate duplicates, update/merge imports, recoverable import batches and export remain pending. Printing is deferred at user request.

## Previous delivery: 0.3.0
Current delivery: `output/Swift_NAVFIT_0.3.0-official-narratives.accdb`, app `0.3.0-prototype`, schema 3. Added an official narrative field to the editor: EVAL block 43, CHIEFEVAL block 40, FITREP block 41. Scroll below revision history to reach the narrative section, which has its own save button (saves the entire draft) and character count. General draft notes stay separate. Changing type before first save changes the narrative label; report type remains fixed after creation.

OfficialComments is Long Text and participates in atomic report/revision/change writes, duplication and restore. Native recordset history writes preserve long multiline content; binary comparisons detect case-only edits. Omitted narrative values in older save calls preserve the stored value, including null versus empty. Historical schema-1/2 snapshots have no official narrative; restoring them clears it to null rather than inventing content.

Validation: 39 narrative assertions across all three forms plus reopen/snapshot persistence, 46 regression checks and 12 direct NAVFIT review checks pass. Populated schema-2-to-3 and schema-3-to-3 migrations compare every original row/field, reopen successfully and preserve source hashes. Legacy draft notes are not automatically converted to official narratives. No full visual layout, Runtime/x86, narrative printed-fit or export compatibility claim is made. Schema-1-to-3 migration has not been tested this milestone.

This is the narrative portion of the official field model. Administrative blocks, traits, recommendations, deliberate import commit, policy validation, routing/comments and NAVFIT export remain pending. Printed layout is deferred at the user's request.

## Previous delivery: 0.2.3
Current delivery: `output/Swift_NAVFIT_0.4.0-narrative-import.accdb`, app `0.4.0-prototype`, schema 4. Deliberate partial NAVFIT import creates a new Swift draft with the recognized form kind and confirmed Comments narrative mapping. The review screen accepts an optional user-entered local member display name and requires a confirmation explaining scope. Dates, traits, folders and other official fields remain unmapped. The editor marks imported drafts accordingly.

Usage: complete workspace setup; Home → Review NAVFIT database → select report → inspect Comments → optionally enter a local member display name → Create draft from narrative → confirm. Open the draft library to continue. It never updates an existing report. Exact same source-file SHA256/report-ID imports are rejected; unknown/null types remain quarantined.

ImportReceipt retains file hash/path, source report ID, author, time, adapter version and value encoding. ImportRawValue retains every data-table field from the source, with table/row/field ordinal, field name, DAO type, null flag and value. Text is kept verbatim; dates are OLE numeric values and scalar values use VBA CStr with decimal mark recorded. These retained values are provenance, not an export adapter. Each receipt currently stores a full source snapshot, which duplicates data across selected-report imports; shared batch storage and a user-facing provenance/history viewer remain pending. Original and staged copies have different hashes and changed files may repeat reports, so cross-file duplicate matching/conflict reconciliation remains pending and is disclosed at confirmation.

Transactions include the new draft, its initial revision, receipt and provenance. Windows native SHA256 matches PowerShell; source is opened read-only and its hash checked before/after capture. A source change rejects the import. No optional backend, PowerShell runtime or external hashing tool is required by the application.

Validation: three typed imports, quarantine, exact-file duplicates, late injected provenance failure with verified rollback, reopen and all 2472 retained values across three imports pass. Also 46 regression checks, 39 narrative assertions and 12 direct-review checks pass. Schema-3-to-4 and schema-4-to-4 upgrades preserve every original row/field including 2472 provenance records; source hashes unchanged; migrated files reopen. All 11 user-data tables in the delivery are empty. Runtime/x86 and complete visual testing remain pending. UI confirmation wiring is checked; the modal confirmation was not automated end to end.

This is narrative-only import, not complete NAVFIT report import/export. Administrative models, field mappings, candidate duplicates, update/merge imports, recoverable import batches and export remain pending. Printing is deferred at user request.

## Previous delivery: 0.3.0
Current delivery: `output/Swift_NAVFIT_0.2.2-tested.accdb`, app `0.2.2-prototype`, schema 2. Added a Home action for read-only NAVFIT staging review, report selection, all-field inspection with DAO type/null indicators, verbatim encoded parents and explicit null display. Review opens the external staging file through read-only DAO without executing its code. No staged data is committed to Swift. Exact ordered field names/types/sizes for all three raw tables are checked before opening the viewer.

Validation: 46 existing regression checks and 12 new review checks pass, including malformed field-name rejection, no draft writes, active-form navigation and unchanged staging hash. VBA compiled in full x64 Access. Full visual inspection and Runtime remain pending. Database schema is unchanged; existing schema-1/2 migration tooling still applies, but migration into this build has not been repeated.

To try it: copy the blank master, complete setup, choose **Review staged NAVFIT**, and select `staged-values.accdb` from the verified batch under `build/navfit-staging-verified4`. The feature is an inspection viewer; staging creation still uses the development script. Report commit and NAVFIT export remain unavailable pending mapping proof.

## Previous delivery: 0.2.1
Current tested delivery: `output/Swift_NAVFIT_0.2.1-tested.accdb`, app `0.2.1-prototype`, Swift schema 2. Earlier artifacts remain preserved and are superseded.

The final suite has 46 passing checks. Installed module exports were compared against source allowing Access identifier case, whitespace and no-argument Call-parenthesis formatting. This check exposed and repaired ambiguous colon labels in refresh actions. Library initial population, editor initial revision population and move-button refresh now have explicit tests. The blank template was verified separately. Runtime, x86 and full visual layout certification remain untested.

Added organizational folder hierarchy, cycle rejection, folder/report moves, contributor move history, Unassigned view, report search and archive filters, contextual help and report counts. These local moves preserve committed report revisions and do not implement official grouping. Help activation and return to the editor were tested against Access's active form, not only function return values. Full visual review on Navy workstations remains pending.

After the workstation crash, the exact candidate passed its complete synthetic suite again. Migration from populated schema 1 and schema 2 test workspaces preserved every original row/field, stable identity, null value, history entry and (for schema 2) folder/move record; originals' hashes remained unchanged and migrated files reopened. The migration helper is removed before publication. It rejects existing destinations, populated targets, unsupported schema versions and unrecognized source data tables. This is development migration tooling requiring full Access, not yet an in-app wizard.

Migration debugging established that native recordset writes avoid the invalid-argument failures encountered when copying nullable/Long Text values through append parameters. The final schema-1 and schema-2 migration tests into 0.2.1 both passed exact comparison and reopen checks. Transactions, exact comparisons and publication only after verification remain enforced. Failed scratch copies are retained separately for diagnosis.

No NAVFIT round-trip compatibility, official block model, policy validation, Runtime/x86 support, tamper resistance or concurrent-use protection is claimed. See `NAVFIT_SCHEMA_FINDINGS.md` for newly observed parent encoding and converter checksum evidence.

## Previous milestone: 0.1

Delivered foundation: `output/Swift_NAVFIT_0.1.0.accdb`, app `0.1.0-prototype`, Swift schema 1. All application objects/data are contained in this file. Source/build/test tools are separate development artifacts.

Implemented: first-run command/profile setup; local library; EVAL/CHIEFEVAL/FITREP type identities; incomplete draft saves; fixed report type after creation; stable GUID; atomic report/revision/field-change writes; no-op save suppression; duplicate; restore as a new revision; archive/recovery and recorded events; support diagnostics; explicit unavailable-feature messaging.

Not implemented: official block-specific/form-specific detail models, policy validation, folders/groups, inheritance/bulk edits, NAVFIT import/export, PDF import, comments/routing, comparison meetings, UTC/offset capture, enforced single-use sessions, backup/restore UI and populated-file migration. The generic draft notes field is deliberately not exported or labeled an official narrative block. This is a foundation prototype, not a release for operational reports.

Validation: VBA compiled and 29 synthetic smoke checks passed in full 64-bit Access 16.0 build 20430, including UI setup/save, empty draft saving, immutable revisions, null dates, identity, archive protections, type protection, a failure injected during snapshot writing with verified rollback, form loading/width bounds and close/reopen persistence. `build/tests/smoke-results.json` records results against the exact delivered build. No visual layout review, Runtime/x86/Navy workstation certification or NAVFIT reimport/print test has been performed.

The delivered database was separately checked through read-only DAO: Workspace, UserProfile, Report, ReportRevision, RevisionValue, RevisionChange and ReportEvent each have zero rows. Test harnesses and synthetic records exist only in disposable test copies. Preserve this blank master and create independent working copies for evaluation.

Build debugging: initial direct creation of unnamed Access forms triggered Save As during failure cleanup. Named template forms and `Quit(acQuitSaveNone)` now avoid this. The first rejection test exposed a VBA debugger dialog; expected failures now run inside a test-only VBA handler and the final suite completes without dialogs. Direct COM varargs calls to parameterized VBA functions were unreliable in this host; smoke tests use the Access expression evaluator. DAO collections are enumerated with bounded indexes to avoid stalled PowerShell COM enumeration.

Compatibility evidence: `TEST_DB.accdb` read-only schema inspected, original SHA256 unchanged. `Reports` has 128 columns, `Folders` 78, `Summary` 6; sample rows cover Eval/Chief/FitRep plus one null ReportType. See `NAVFIT_SCHEMA_FINDINGS.md`. No source report values were copied into the prototype. No converter code was executed. The upgrade tool was locked by another process; checksum/code inspection is unverified. Exact NAVFIT executable/build, genuinely blank template, numeric code interpretation, logical folder joins and all-field round-trip/printing evidence remain needed.

Next milestone: resolve those P1 mapping questions and establish a minimal isolated compatibility harness before extending the official editor. The supplied populated test file is not a verified blank export template.
