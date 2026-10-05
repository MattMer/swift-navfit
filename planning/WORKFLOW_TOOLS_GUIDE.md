# Workflow tools — Swift NAVFIT 0.6.0

> Current native navigation: use the [workspace guide](UX_WORKSPACE_GUIDE.md) and
> [current state](CURRENT_STATE.md). This guide retains service details and
> earlier verification history; legacy screen instructions may differ.

Current blank application: `output/Swift_NAVFIT_0.6.0-workflow-tools.accdb`, schema 6. Copy the master for use and complete first-run setup. Existing workspaces need migration into a new file; `scripts/upgrade-workspace.ps1` defaults to this template. Never open an older populated workspace expecting new application objects to appear automatically.

## Saved draft checklist

Select a report in the library and choose **Draft checklist**, or use **Check latest saved draft** near the bottom of the editor. Findings cover local title/member display, period dates, reversed date order, blank narrative, partial imports, pending recovery checkpoints and archive state. Select a finding for guidance. Refresh reads the latest saved values; unsaved editor entries are excluded and preserved. Incomplete drafts remain saveable.

This is a local completeness aid. It does not assess Navy policy, complete official fields, grouping, recommendation limits or export readiness. General notes do not satisfy the narrative prompt. Narrative blankness is a drafting prompt rather than a determination of an official violation.

## Incoming narrative reconciliation

From Home, choose **Review NAVFIT database**. Select the source report, choose a destination saved draft in the bottom dropdown, then **Compare with selected saved draft**. The next screen displays both narratives. **Keep Swift value / close** leaves the report unchanged. **Accept incoming narrative** requires confirmation and replaces only that saved field, retaining other Swift fields and adding a revision and source receipt/provenance.

Close the editor before acceptance. The service rejects a changed source hash, a changed destination revision, differing form types, archived/locked reports and exact repeat imports. Mapping is limited to Reports.Comments and form kind. This is explicit two-way reconciliation, not identity matching or a three-way returned-export merge. Dates, traits, recommendations and source folder links remain unmapped. Exact source hash/report ID duplicate protection does not detect duplicates across altered or differently packaged files.

## Narrative snippets

Open **Narrative snippets** from Home or the editor. New snippets require a title, body, attribution and applicable report type (or ALL). Search covers title and body. Select an existing snippet to view/edit it; use New snippet for a separate entry.

With a draft editor open, **Append to open draft** requires confirmation, preserves existing narrative text and inserts a line break before the snippet. Type mismatches are rejected. Insertion updates pending editor text and requests a recovery checkpoint; it does not commit a revision until Save draft. Snippets are locally curated, not automatically policy-approved, and should be reviewed for the particular member. Snippet attribution stays in the library; a separate report-to-snippet insertion audit is not implemented.

## Saved library column views

The library dropdown offers **Drafting**, **Dates** and **History**. The selected preset persists for the active contributor profile. All retain report ID selection, type, title and member display. Dates shows period boundaries; History shows modified time and current revision ID. Search, folder and archive filters continue to apply. Arbitrary column selection/reordering and multiple named custom presets are later extensions.

## Bulk folder moves

From the library choose **Bulk moves**. Select multiple active drafts, choose a destination (including Unassigned), and preview. The preview identifies selected reports and current folders. Applying requires confirmation and rechecks the selection, destination, current revisions and latest move IDs. The operation is atomic and recorded as a batch. Organizational folder assignments change; report revisions do not.

This is the first bulk-change operation. Bulk official-field editing, inheritance and official summary-group effects remain pending.

## Operation history and guarded undo

Home → **Operations and import undo** shows Started, Completed, Failed and Reviewed entries. Select an operation for its outcome. Started entries may represent interrupted work; inspect receipts and saved data before retrying. **Acknowledge selected incomplete operation** marks the entry Reviewed without retrying or changing report data. Successfully committed imports finish their journal entry inside the same transaction.

Select an import receipt and choose **Undo selected import**. New imported drafts are archived; reconciled narratives restore the preceding revision through a new compensating revision. Provenance and historical revisions remain. Later edits, archive/lock state or missing applied-revision metadata block undo. Receipts from earlier app versions lack the new metadata and cannot be automatically undone.

Select a bulk-folder operation and choose **Undo selected folder operation**. Preflight checks every affected report. Any later report revision or folder move blocks the entire undo. Successful undo records compensating moves and marks the batch rows undone. A forced late batch failure was tested to leave no partial moves.

Journal coverage currently includes narrative imports/reconciliation, import undo, bulk folder moves and their undo. Export is unavailable; development upgrade tooling retains its separate migration evidence. Automatic resumable jobs are not implemented.

## Verification and remaining limits

54 integrated tools assertions passed, including exact reconciliation/undo narrative values, unchanged other fields, checklist behavior, pending editor preservation, snippet search/append/type guards, persisted views, duplicate and stale-source guards, later-edit undo guards and injected batch rollback. Existing suites passed: 46 draft regression checks, 39 narrative assertions, 22 comparison assertions, 34 recovery assertions plus timer capture, 12 direct source-review checks, and complete import/provenance tests.

Schema-5-to-6 and schema-6-to-6 migration compare every original row/field, preserve source hashes and reopen targets. The populated schema-6 test includes 2 snippets, 1 preference, 14 journal entries, 4 batch-change rows and 2472 provenance values. All 16 user-data tables are empty in the delivered master. The delivery is byte-identical to the tested candidate, as confirmed by SHA256 in the manifest.

Full visual inspection, Runtime/x86, older-schema-to-6 migration and general failure recovery across power loss remain pending. Printed layout is deferred. NAVFIT export and complete report import remain unavailable.
