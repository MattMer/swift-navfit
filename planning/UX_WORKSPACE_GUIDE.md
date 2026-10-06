# Swift NAVFIT native workspace guide

Current prerelease: **0.24.0-preview, schema 23**, based on development r75. The distributed 0.21.0
checkpoint is older. Start at [CURRENT_STATE.md](CURRENT_STATE.md) for current
artifacts, verification evidence and remaining work. Earlier feature guides
are historical service documentation; use this guide for current navigation.

## Start here

Use the synthetic demo to explore. Use the blank candidate for a separate workspace; complete the command and contributor setup. Keep existing workspaces intact. Reports, Summary groups and My Work are the main destinations. Local contributor profiles provide attribution, not authentication.

## Create, edit and save a report

Reports > New report: choose EVAL, CHIEFEVAL or FITREP. Edit sections in official block order. Checkbox blocks remain explicit selections; SSNs retain leading zeros. Structured name parts stay separate from stored full names. Save draft creates an atomic revision. Section changes preserve pending values without creating revisions. Leaving a changed report saves automatically. If saving fails, correct the inline errors; pending values remain. Draft notes are local, separate from official comments.

## Find reports and history

Reports provides search, form-type, folder, review and archive filters, saved views and selection actions. Folders organize reports independently of summary groups. Open a report to inspect History, compare revisions or restore a previous revision as a new one. Duplicate creates a separate draft. Archived reports can be recovered; local approval must be reopened through Review before editing.

## Summary groups and comparisons

Summary groups lists groups with report counts. Open a group for Members, Shared defaults, Review, Comparison and History. Rename, archive and recover act on the group; reports remain independent. Bulk changes and defaults use the same official sections as a report. Mixed, unchanged, staged and cleared values are distinct. Preview full changes, then Apply. Comparison uses editable trait cells and staged proposals; an internal baseline is captured automatically. Changed membership, revisions or locks require a fresh preview. Open All summary-group validation issues for member findings, grouping differences and provisional EP/MP limits with policy references.

## Reporting seniors and averages

Resources > Reporting seniors and RSCA history stores reusable senior details and manually entered rank-specific historical baselines. Apply details deliberately; directory edits do not silently update reports. Group > Members > Reporting senior and PSG RSCA shows saved sums, counts, exclusions, baseline date/source and exact versus estimated calculations. Confirm the historical baseline excludes the current group. Report > Review > Calculation details separates pending and saved values. Arithmetic and manual provenance do not certify Navy policy compliance.

## My Work and local review

My Work combines assigned reviews, current ordered-route steps, ready local approvals and related unresolved comment threads. Each row names the member, report period, contributor, status and next action. Assignment due dates are shown when recorded; routes and comments have no invented due dates or assignees. Future route steps stay out of the current contributor's queue. Requesters can see stale routes that need restarting. Select a row and choose Open selected work to reach that route or thread directly, or Open report to edit the report. Accept/Return apply only to an eligible selected assignment; stale and locked decisions are disabled. Completed work and all contributors are available through the history filter.

## Checks, snippets and imports

The report action-area validation button and Report > Review > Validation issues check pending values live. Summary group > All summary-group validation issues includes all members and pending edits from the open member. Select a finding for its instruction, enclosure, printed page, paragraph/table and excerpt, then open the affected report block. Invalid entries, incomplete drafts, policy issues and Needs review are distinct. Storage checks are application rules. Group capacities are provisional until eligible membership is verified; no manual roster freeze is required. Live checks create no saved revisions or validation runs. Coverage is limited, including manual review of LDO/CWO categories, UIC exceptions and earlier/future instruction applicability. Reports > Draft checklist retains the saved-value checklist. Resources provides attributed narrative snippets. Reports > Import reports offers NAVFIT report or structured PDF import: Choose, Preview, Apply, Results. PDF mapping covers 40 EVAL, 51 CHIEFEVAL and 36 FITREP saved fields. Leave the merge target blank to create a new draft, or choose an existing draft and Fill empty fields / Replace mapped values. Full before/after previews show replacements and retained values. Blank or missing PDF fields never clear existing values. Close the target editor and resolve pending recovery before merging. Changed or locked targets require a fresh preview. Settings & Help > Operation history can undo an import while no later edit exists. Use Check PDF extraction on this computer after choosing a PDF; desktop Adobe Acrobat is required, Reader/browser viewers are insufficient. Changed sources or mapping choices require a fresh preview. Results can open the imported or merged report. Source field inspection and existing-draft reconciliation remain available. PDF imports use SSN by default; legacy DoDID mapping requires explicit compatibility selection. All imports retain provenance and are limited in scope. NAVFIT import maps 74 saved report values and retains all 128 native report columns, folders and summary values. Choose a counseling-date century explicitly when needed. NAVFIT exchange exports one saved report with a separate hashed Swift baseline. Keep that baseline locally and select it when the edited NAVFIT file returns. Preview all baseline/current/returned values and resolve every editable-field conflict before Apply; Apply also retains the displayed native-only source changes. Source changes, later revisions, pending recovery, archive and local approval guards apply. Summary group > Export group to NAVFIT98A previews the saved roster and promotion totals, then exports one native folder with all selected members. Keep the frozen group baseline locally. Imported native-only report values are preserved; new drafts have no native RSCA unless Write exact PSG RSCA is selected. That option requires complete traits and an exact historical sum. Group-return reconciliation, arbitrary-source matching and unrecognized counseling encodings remain unsupported; NOT REQ and NOT PERF are preserved. Open the exported file in NAVFIT98A, validate reports, and use its report/folder/summary printing actions. Swift does not yet provide the native print engine directly. Counseling accepts a real ISO date, NOT REQ, or NOT PERF; NOT PERF requires an explanation. Live checks include counseling period, required career recommendation, partial-NOB trait limits, adverse-mark justification reminders and Regular reports exceeding 15 months (approval must be verified). Resources > NavyWriter catalog and Narrative snippets both filter by rate. Legacy untagged snippets remain Unassigned; assign a rate when editing. Collection links open the complete source; only reviewed short excerpts can be imported.

## Recovery, backups and migration

Overview or Settings & Help > Recover pending edits opens checkpoints, separate from saved revisions. If the saved baseline changed, recover as a new draft. Settings & Help has backup/restore and operation history with guarded undo. Close the source before backup or migration. Transfer existing data into a separate blank candidate using upgrade-workspace.ps1 with an explicit TemplatePath. Never migrate into the seeded demo. The source stays unchanged; a new destination and migration evidence are created.

## Preview limits and diagnostics

This is version 0.23.0-preview, schema 21, a native Access review candidate. Automated native checks do not establish screen-reader or physical display/DPI acceptance. Desktop Access is verified within the documented tests; Runtime and x86 remain unverified. No network integration, synchronization, authentication or official signatures are added. Settings & Help > Support and limitations provides technical diagnostics. Read CURRENT_STATE.md for exact artifact evidence and outstanding work.

## Feature destinations

| Existing capability | Current destination |
|---|---|
| Official fields, identifiers, structured names, contacts, traits, narratives | Report sections in official order; local title/name/notes in Draft notes |
| Save, checkpoints, entry errors, completeness | Report action area and Review > Draft checklist; Overview recovery |
| Revisions, compare/restore, duplicate, archive/recover | Report History and Reports selection actions |
| Folders, search, filters, saved views, bulk folder moves | Reports; operation history for guarded undo |
| Group lifecycle, membership, defaults, shared blocks | Summary groups list and focused group Members/Shared defaults |
| Meeting proposals, editable traits, pending group average | Focused group Comparison, Preview and Apply |
| Senior directory, rank RSCA baselines and PSG RSCA | Resources > Reporting seniors; group Members calculator |
| Assignments, due dates, contributor switching | My Work and Settings & Help |
| Comments/replies, routes, saved-revision review, local approval/reopen | Report or focused group Review |
| Live report/group validation and policy excerpts | Report action-area issue button; focused group validation issues |
| Limited versioned Navy checks | Review > Run checks; source/effective-period scope remains limited |
| Snippets and reference vocabulary | Resources, report narrative insertion and trait guidance |
| NAVFIT/PDF review, supported imports, reconciliation, provenance | Reports import tools and import review/provenance |
| Audited batches, undo and operation history | Contextual group/report tools; Settings & Help > Operation history |
| Backup/restore, workspace transfer and diagnostics | Settings & Help and the candidate migration script |
| Legacy DoDID values | Retained for compatibility; normal report editing uses SSN |

No saved feature data is removed by navigation changes. Printing, full NAVFIT
export and broader interoperability remain deferred/incomplete, rather than
being treated as completed destinations.

Focused saved-report review and Reports viewport return are documented in [UX_REVIEW_RETURN_2026-10-04.md](UX_REVIEW_RETURN_2026-10-04.md).
