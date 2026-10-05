# Swift NAVFIT native workspace guide

Current native preview: **0.22.0-ux-preview, schema 21**. The distributed 0.21.0
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

Summary groups lists groups with report counts. Open a group for Members, Shared defaults, Review, Comparison and History. Rename, archive and recover act on the group; reports remain independent. Bulk changes and defaults use the same official sections as a report. Mixed, unchanged, staged and cleared values are distinct. Preview full changes, then Apply. Comparison uses editable trait cells and staged proposals; an internal baseline is captured automatically. Changed membership, revisions or locks require a fresh preview.

## Reporting seniors and averages

Resources > Reporting seniors and RSCA history stores reusable senior details and manually entered rank-specific historical baselines. Apply details deliberately; directory edits do not silently update reports. Group > Members > Reporting senior and PSG RSCA shows saved sums, counts, exclusions, baseline date/source and exact versus estimated calculations. Confirm the historical baseline excludes the current group. Report > Review > Calculation details separates pending and saved values. Arithmetic and manual provenance do not certify Navy policy compliance.

## My Work and local review

My Work combines assigned reviews, current ordered-route steps, ready local approvals and related unresolved comment threads. Each row names the member, report period, contributor, status and next action. Assignment due dates are shown when recorded; routes and comments have no invented due dates or assignees. Future route steps stay out of the current contributor's queue. Requesters can see stale routes that need restarting. Select a row and choose Open selected work to reach that route or thread directly, or Open report to edit the report. Accept/Return apply only to an eligible selected assignment; stale and locked decisions are disabled. Completed work and all contributors are available through the history filter.

Report > Review opens routes and comments in the report context. Review operates on saved revisions. With pending edits choose Save and continue, Use saved version or Cancel. Stale routes need a restart against the current revision. Local approval is a local workflow lock, not an official signature. Review comments retain replies and resolution history. Resolved threads leave the pending queue; reopening a thread returns it to the relevant contributors' queues.

## Checks, snippets and imports

Report > Review > Draft checklist uses saved values and jumps to the relevant official section. Inline entry errors, incomplete draft information and limited Navy checks are separate. Navy checks cover selected source-linked rules and their effective report-period range. Resources provides attributed narrative snippets. Reports > Import reports offers NAVFIT narrative or structured PDF import: Choose, Preview, Apply, Results. Full before/after mappings show what will be created. Changed sources or mapping choices require a fresh preview. Results can open the created report. Source field inspection and existing-draft reconciliation remain available. PDF imports use SSN by default; legacy DoDID mapping requires explicit compatibility selection. All imports retain provenance and are limited in scope. Printing and full NAVFIT export remain deferred/incomplete.

## Recovery, backups and migration

Overview or Settings & Help > Recover pending edits opens checkpoints, separate from saved revisions. If the saved baseline changed, recover as a new draft. Settings & Help has backup/restore and operation history with guarded undo. Close the source before backup or migration. Transfer existing data into a separate blank candidate using upgrade-workspace.ps1 with an explicit TemplatePath. Never migrate into the seeded demo. The source stays unchanged; a new destination and migration evidence are created.

## Preview limits and diagnostics

This is version 0.22.0-ux-preview, schema 21, a native Access review candidate. Automated native checks do not establish screen-reader or physical display/DPI acceptance. Desktop Access is verified within the documented tests; Runtime and x86 remain unverified. No network integration, synchronization, authentication or official signatures are added. Settings & Help > Support and limitations provides technical diagnostics. Read CURRENT_STATE.md for exact artifact evidence and outstanding work.

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
