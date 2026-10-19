# Swift NAVFIT

**Draft, review, and manage Navy evaluations in one local workspace.**

Swift NAVFIT is a Windows application built in Microsoft Access for EVALs, CHIEFEVALs, and FITREPs. It brings report drafting, summary groups, member history, local review, revision tracking, and recovery into one `.accdb` file.

The goal is to reduce the spreadsheets, emailed copies, repeated data entry, and manual tracking that surround the report-writing process. Swift works alongside NAVFIT98A through supported import/export workflows; NAVFIT98A remains the destination for final native validation and printing.

## What Swift adds to the workflow

### Know what changed—and recover it

- **Field-level revision history:** see who saved a change, when it changed, and the previous/new values.
- **Compare and restore revisions:** inspect differences and restore an earlier version as a new revision without erasing history.
- **Unsaved-edit recovery:** automatic/manual checkpoints preserve pending editor work separately from saved revisions.
- **Clear save feedback:** distinguish unsaved edits, recovery checkpoints, failed saves, and committed revisions. Failed saves retain pending values for correction and retry.
- **Archive and recover:** keep reports out of active views without losing their history.

### Coordinate review without a separate tracker

- **My Work:** one queue for assigned reviews, current route steps, ready local approvals, and unresolved comments, with status and next actions.
- **Ordered review routes:** assign contributors to steps, accept or return work with decision notes, and retain earlier decisions when a stale route is restarted.
- **Revision-linked comments:** reply, resolve, and reopen discussion threads tied to saved report values.
- **Stale-review protection:** later report changes invalidate earlier review eligibility instead of silently carrying approval forward.
- **Local attribution:** contributor names, decision notes, and timestamps make the local review trail readable.

### Work across a command and summary groups

- **Command member directory:** reuse identities and filter report creation by paygrade or rank/rate; create reports individually or in batches.
- **Prior evaluation history:** keep historical evaluations alongside each member for reference without replacing the current report.
- **Shared fields and defaults:** apply group or command defaults with a visible scope, per-field confirmation, and disclosure when customized values or local approvals are affected.
- **Live member tables:** edit trait cells in place; failed cells remain available for retry or deliberate discard.
- **Comparison and proposals:** inspect group marks and stage changes with safeguards against changed membership, revisions, and locks.
- **Reporting-senior directory and RSCA history:** reuse senior details and dated baselines; distinguish exact historical totals from rounded-average estimates and select a group's baseline deliberately.

### Catch problems while drafting

- **Live, cited findings:** check pending report values and groups, with references and navigation to the affected block.
- **Distinct issue types:** separate invalid entries, incomplete drafts, policy issues, and matters needing human review.
- **Drafts remain editable:** incomplete information does not prevent an otherwise valid draft save.
- **Guided report pages:** fixed Front/Back layouts, all seven traits together on Front Page, contextual help, and consistent navigation/status colors.
- **Narrative tools:** rate-filtered, attributed snippets and local draft notes remain separate from the official narrative.

### Bring existing data forward

- **NAVFIT98A exchange:** import supported report/group data and export saved reports or summary groups while retaining native provenance.
- **Conflict-aware report returns:** compare the exported baseline, current Swift report, and returned NAVFIT file before applying conflicting edits.
- **Structured PDF import:** preview mapped values and create a new draft or deliberately fill/replace an existing draft. Supported extraction requires desktop Adobe Acrobat.
- **Previous-Swift workspace transfer:** move supported older workspace data and history into a separate blank build, preserving the source and verifying the copied values.
- **Backup/restore and operation history:** guarded tools for protecting the workspace and undoing supported operations.

## Swift NAVFIT compared with NAVFIT98A

NAVFIT98A already provides report editing, folder/default management, validation, import/export, automatic summary functions, and report/folder/summary printing. Those are **shared foundations**, not features invented by Swift. Swift's main additions are the history, recovery, and review services around the report.

| Capability | NAVFIT98A's documented workflow | Swift NAVFIT |
| --- | --- | --- |
| Saved-change history | Edit and save reports; field-level revision comparison/restore is not described in the reviewed guides. | Contributor-attributed revisions, before/after values, comparison, and restore as a new revision. |
| Pending-edit checkpoints | A separate unsaved-checkpoint recovery workflow is not described. | Recover pending work independently of committed revisions; changed baselines can recover as a new draft. |
| Review tracking | Ordered contributor routes and a unified review-work queue are not described. | My Work, assignments, ordered steps, decision notes, stale-route detection, and local approval. |
| Discussion history | Revision-linked discussion threads are not described. | Comments, replies, resolve/reopen history, and direct navigation from My Work. |
| Member reuse and history | Report/folder-centered editing and import/export. | A command directory, rank/paygrade filters, batch report creation, and separate prior evaluations. |
| Reporting-senior averages | Summary/average calculations are part of NAVFIT98A. | Reusable senior records and dated RSCA baselines, with exact-versus-estimated provenance and deliberate group baseline selection. |
| Shared group values | Folder defaults and common report values are supported. | Explicit affected-report scope, confirmation of replacements, and reopening of affected local approvals. |
| Validation | Native individual-report and folder validation. | Live checks of pending values, issue categories, citations, and links to report fields; coverage is bounded. Native validation is still required after export. |
| Returned report files | Report/folder import and export. | A retained export baseline and three-way conflict preview for supported single-report returns. |
| Final native printing | Built-in report, folder, and summary printing. | Export to NAVFIT98A; Swift does not yet provide an equivalent native print engine or verified final printed fit. |

The NAVFIT98A column is based on the [Navy user guide](https://www.mynavyhr.navy.mil/Portals/55/Career/PerformanceEvaluation/NAVFIT98A%20Version%2030%20user%20guide.pdf), [v33 FAQ](https://www.mynavyhr.navy.mil/Portals/55/Career/PerformanceEvaluation/NAVFIT98A%20v2.2.0.33%20Frequently%20Asked%20Question.pdf), and the locally inspected v33 Patch 2 package. “Not described” identifies an addition to the reviewed workflow, not an exhaustive claim about every NAVFIT98A version. See the [package findings](planning/NAVFIT_PACKAGE_FINDINGS.md), [native validation comparison](planning/NAVFIT_VALIDATION_COMPARISON_2026-10-14.md), and [print-engine findings](planning/NAVFIT_PRINT_ENGINE_FINDINGS.md) for the evidence and version boundaries. Historical comparisons may describe gaps addressed by later Swift work.

## Try Swift NAVFIT

**Requires Windows and full desktop Microsoft Access. Free Access Runtime is not a supported target.**

1. Choose a blank workspace for a new command, or a synthetic demo for practice, from the [available releases](https://github.com/MattMer/swift-navfit/releases).
2. Keep a working copy and open it in desktop Access. Complete command and contributor setup.
3. Create or select a summary group, add members/reports, and use My Work for local review. The top-right **?** opens help for the current window.
4. To carry an older Swift workspace forward, open a **new blank build** and choose **Import data from a previous Swift NAVFIT version**. Preview the source, import, then close and reopen the destination. Never overwrite an existing workspace with a blank distribution.

**Source and downloads can differ.** The repository currently develops `0.25.0-dev` / schema 25. Older release downloads may not include every feature or recent refinement described here. New source commits do not automatically publish a new Access binary.

The latest local demo adds a group with **zero member-report findings** and another with **three deliberate report issues**. Both retain roster-completeness and provisional-capacity reminders. Recent development also adds wider Front/Back viewing areas, contextual help, consistent colors, and faster page switching; see [recent changes](planning/RECENT_DEVELOPMENT.md).

## Important boundaries

- This is an independent development project, not an official Navy system or a replacement for governing instructions and human review.
- **Local approval is not official submission, signature, or Navy approval.** Group workflow status is also separate from member-report review.
- Contributor profiles provide attribution, not authentication. Use one authoritative workspace file, one contributor at a time. Separate copies do not synchronize.
- Validation and exchange cover implemented cases, not every policy or native encoding. Returned-group reconciliation, arbitrary-source matching, and full native printing remain outside the supported workflow.
- Save checkpoints, revisions, backups, and migration have different purposes. A workspace transfer preserves a source and creates a separate destination; it does not merge independent populated files.

## Build and contribute

Editable source lives in `src/`; Access builders and focused checks live in `scripts/` and `tests/`. Databases, supplied Navy files, and generated scratch builds stay out of Git. Build into a fresh filename on Windows with desktop Access:

```powershell
powershell -NoProfile -File scripts/build.ps1 -OutputPath output/Swift_NAVFIT_review.accdb
powershell -NoProfile -File scripts/smoke-test.ps1 -BuildPath output/Swift_NAVFIT_review.accdb
```

See [contributing](CONTRIBUTING.md) for prerequisites and checks. New releases must pass the retained-schema data-transfer/rollback checks and blank/demo inventory gates. Physical scaling, accessibility, and final printed fit require their own acceptance; automated checks do not establish them.

## Guides and project status

- [Recent development and verification](planning/RECENT_DEVELOPMENT.md)
- [Current state](planning/CURRENT_STATE.md) and [roadmap](planning/ROADMAP.md)
- [Summary groups, review routes, and workflow foundations](planning/MAJOR_FOUNDATIONS_GUIDE.md)
- [Workspace guide](planning/UX_WORKSPACE_GUIDE.md)
- [Review assignments and My Work](planning/REVIEW_ROUTING_GUIDE.md)
- [Review comments](planning/REVIEW_COMMENTS_GUIDE.md)
- [Reporting seniors and averages](planning/REPORTING_SENIOR_GUIDE.md)
- [Workspace transfer and fixed report pages](planning/WORKSPACE_COMPAT_PAGES_R115_2026-10-17.md)
- [Import provenance](planning/PROVENANCE_GUIDE.md)
- [Build and verification history](planning/BUILD_STATUS.md)

This private repository has no selected open-source license. Coordinate use and redistribution with its owner.
