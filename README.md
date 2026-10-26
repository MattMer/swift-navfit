# Swift NAVFIT

**Draft, review, and manage Navy evaluations in one local workspace.**

[User guide (PDF)](docs/Swift_NAVFIT_User_Guide.pdf) | [Read the guide online](docs/Swift_NAVFIT_User_Guide.md)

The 22-page guide covers setup, drafting, notes and replies, validation, summary groups, RSCA, local review, exchange, recovery, and upgrades. It follows the procedure-based organization of the eNavFit guide and describes Swift's current local workflow.

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

- **My Work:** one queue for assigned reviews, current route steps, and ready local approvals, with status and next actions. A separate Report notes tab opens shared conversations.
- **Ordered review routes:** assign contributors to steps, accept or return work with decision notes, and retain earlier decisions when a stale route is restarted.
- **Shared report conversations:** notes and replies live inside the report workspace, independent of revisions and approval. Private unfinished drafts can be saved and resumed; conversations need no resolve/reopen workflow.
- **Stale-review protection:** later report changes invalidate earlier review eligibility instead of silently carrying approval forward.
- **Local attribution:** contributor names, decision notes, and timestamps make the local review trail readable.

### Work across a command and summary groups

- **Command member directory:** reuse identities and filter report creation by paygrade or rank/rate; create reports individually or in batches. [Review possible duplicate Sailors](docs/Duplicate_Sailor_Review.md), inspect masked identities and affected links, and merge with retained history and guarded undo.
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

- **NAVFIT98A exchange:** import supported report/group data and export all saved summary-group members using the internal native template. Select a group, choose an export directory, and export. Incomplete drafts, validation findings and archived/approved status do not block export. Pending edits remain separate; export does not imply approval or submission.
- **Conflict-aware report returns:** compare the exported baseline, current Swift report, and returned NAVFIT file before applying conflicting edits.
- **Structured PDF import:** preview mapped values and create a new draft or deliberately fill/replace an existing draft. Extraction runs locally without Acrobat or an internet connection.
- **Previous-Swift workspace transfer:** bulk-copy supported older workspace data/history into a separate blank build, with exact database-engine verification and responsive progress between batches. Source and rollback protections remain.
- **Backup/restore and operation history:** guarded tools for protecting the workspace and undoing supported operations.

## Swift NAVFIT compared with NAVFIT98A

NAVFIT98A already provides report editing, folder/default management, validation, import/export, automatic summary functions, and report/folder/summary printing. Those are **shared foundations**, not features invented by Swift. Swift's main additions are the history, recovery, and review services around the report.

| Capability | NAVFIT98A's documented workflow | Swift NAVFIT |
| --- | --- | --- |
| Saved-change history | - | Contributor-attributed revisions, before/after values, comparison, and restore as a new revision. |
| Pending-edit checkpoints | - | Recover pending work independently of committed revisions; changed baselines can recover as a new draft. |
| Review tracking | - | My Work, assignments, ordered steps, decision notes, stale-route detection, and local approval. |
| Discussion history | - | Unified report notes and replies, private unfinished drafts, and direct navigation from My Work; independent of approval. |
| Member reuse and history | - | A command directory, rank/paygrade filters, batch report creation, separate prior evaluations, and reviewed duplicate merging with retained history and guarded undo. |
| Reporting-senior averages | Summary/average calculations are part of NAVFIT98A. | Reusable senior records and dated RSCA baselines, with exact-versus-estimated provenance and deliberate group baseline selection. |
| Shared group values | Folder defaults and common report values are supported. | Explicit affected-report scope, confirmation of replacements, and reopening of affected local approvals. |
| Validation | Native individual-report and folder validation. | Embedded checks plus a modeless Validation Issues popout, pending values, issue categories, citations, and links to report fields; coverage is bounded. Native validation is still required after export. |
| Three-way returned-report reconciliation | - | A retained export baseline and three-way conflict preview for supported single-report returns. |
| Final native printing | Built-in report, folder, and summary printing. | Export to NAVFIT98A; Swift does not yet provide an equivalent native print engine or verified final printed fit. |

The NAVFIT98A column is based on the [Navy user guide](https://www.mynavyhr.navy.mil/Portals/55/Career/PerformanceEvaluation/NAVFIT98A%20Version%2030%20user%20guide.pdf), [v33 FAQ](https://www.mynavyhr.navy.mil/Portals/55/Career/PerformanceEvaluation/NAVFIT98A%20v2.2.0.33%20Frequently%20Asked%20Question.pdf), and the locally inspected v33 Patch 2 package. A **-** means no equivalent feature was identified in the reviewed NAVFIT98A documentation/package; it is not an exhaustive claim about every version. NAVFIT98A still supports ordinary report/folder import and export. See the [package findings](planning/NAVFIT_PACKAGE_FINDINGS.md), [native validation comparison](planning/NAVFIT_VALIDATION_COMPARISON_2026-10-14.md), and [print-engine findings](planning/NAVFIT_PRINT_ENGINE_FINDINGS.md) for the evidence and version boundaries. Historical comparisons may describe gaps addressed by later Swift work.

## Data security and offline operation

Swift NAVFIT stores reports, revisions, review history, and recovery checkpoints locally in the workspace `.accdb`. Normal drafting and review do not communicate with the internet: the app does not upload report data, synchronize with an online service, or send telemetry. No internet account, cloud service, or separate application-specific network protection is required.

The workspace uses the Windows file/folder access controls and storage protections of its location. Place it in the same approved, protected folder you would use for other files containing the same data; it inherits that folder's permissions under normal Windows inheritance settings and respects explicit file permissions. Existing device encryption, backup, and organizational data-handling rules continue to apply. Swift does not add a separate authentication or encryption layer, and contributor profiles are attribution labels rather than access controls. SSNs remain readable inside the database, including revisions, recovery and retained import data; masked displays do not encrypt storage.

Optional source links, such as NavyWriter, open an external website only when selected; report contents are not included in those links. Microsoft Access/Office and any cloud synchronization configured for the containing folder operate under their own settings, independently of Swift.

## Architecture, security and accessibility documentation

For command reviewers, the ISSM and anyone assessing Swift NAVFIT for use:

- **[Architecture description](docs/architecture/Swift_NAVFIT_Architecture_Description.md)**, aligned with the DoD Architecture Framework (DoDAF) 2.02 and structured per ISO/IEC/IEEE 42010:
  - stakeholders, concerns and the viewpoints chosen for a standalone desktop application, with reasons for the ones left out;
  - 14 fit-for-purpose views: overview, operational concept, activities, rules, state transitions, data models, interfaces, functions, measures, schema evolution and standards.

  Five of the views are generated from the source and a fresh build, and a test fails if any of them drifts. The description is DoDAF-*aligned*, not certified conformant: formal conformance requires DoD architecture tooling (DM2/PES exchange).
- **[DevSecOps alignment matrix](docs/architecture/DevSecOps_Alignment.md):** every activity in the DoD Enterprise DevSecOps Guidebook, marked as done (with evidence), partial, command-owned or not applicable. No gaps remain.
- **[Software bill of materials](docs/architecture/SBOM.md)** (also in [CycloneDX](docs/architecture/sbom.cdx.json)), generated from the build. Swift uses no third-party runtime libraries; everything it calls is part of Windows or Microsoft Office.
- **[Threat model](docs/architecture/Threat_Model.md):** STRIDE over Swift's assets, trust boundaries and file interfaces, the mitigations in the code, recommendations, and the code-signing plan. Every importer is tested against 55 generated malformed and control inputs.
- **Section 508 accessibility:**
  - automated checks for contrast, screen-reader names and Windows contrast themes;
  - the [screen-reader test procedure](docs/Accessibility_Screen_Reader_Test.md);
  - [user guide section 3.4](docs/Swift_NAVFIT_User_Guide.md#34-accessibility).

## Try Swift NAVFIT

**Requires Windows and full desktop Microsoft Access. Free Access Runtime is not a supported target.**

### Current release: 0.49.0

[**Swift NAVFIT 0.49.0**](https://github.com/MattMer/swift-navfit/releases/tag/v0.49.0) (schema 35) passed the full release gate on 26 October 2026. The release page lists each file's SHA-256 hash.

- **Swift_NAVFIT_Blank_0.49.0.accdb:** an empty workspace for a command. Start real work here.
- **Swift_NAVFIT_Demo_0.49.0.accdb:** synthetic members, reports and groups for practice. Never put real data in it.

> **Four acceptance checks are still untested:**
> 1. printed fit against the official forms;
> 2. screen-reader (NVDA/JAWS) use;
> 3. a workstation without Adobe Acrobat;
> 4. a Navy-managed workstation under its normal Office policy.
>
> The files are also unsigned: choose **Enable Content** when Access asks.

To start:

1. Download the Blank (or the Demo to practice), keep the original, and open a copy in desktop Access. Choose **Enable Content**.
2. Complete command and contributor setup. Swift asks you to [choose your contributor at every launch](docs/Contributor_Startup.md).
3. Create or select a summary group, add members and reports, and use My Work for local review. The top-right **?** opens help for the current window, and the [user guide](docs/Swift_NAVFIT_User_Guide.md) covers every procedure.
4. **Upgrading from an earlier Swift workspace:**
   1. Open a copy of the new Blank.
   2. Choose **Settings & Help → Import data from a previous Swift NAVFIT version**. Every earlier schema imports, and every value is verified before commit.
   3. Close and reopen the workspace.

   Never overwrite an existing workspace with a Blank.

### Development pre-releases

Releases tagged `-dev` on the [releases page](https://github.com/MattMer/swift-navfit/releases) are test builds of work in progress. They aren't verified releases.

### Earlier releases

r128 (app 0.27.1-dev, schema 27) remains available: [Blank](https://drive.google.com/file/d/17YWFTI4N2ZbDG04K-qAddWQz7bEM1kjO/view?usp=drivesdk) and [synthetic Demo](https://drive.google.com/file/d/10NEFdhRhJQ9xC_baxNifrRgs4SR4noUg/view?usp=drivesdk). Its workspaces import into 0.49.0.

## Important boundaries

- This is an independent development project, not an official Navy system or a replacement for governing instructions and human review.
- **Local approval is not official submission, signature, or Navy approval.** Group workflow status is also separate from member-report review.
- Contributor profiles provide attribution, not authentication. Use one authoritative workspace file, one contributor at a time. Separate copies do not synchronize.
- Validation and exchange cover implemented cases, not every policy or native encoding. Returned-group reconciliation, arbitrary-source matching, and full native printing remain outside the supported workflow.
- Save checkpoints, revisions, backups, and migration have different purposes. A workspace transfer preserves a source and creates a separate destination; it does not merge independent populated files.

## Build and contribute

Editable source lives in `src/`; Access builders and focused checks live in `scripts/` and `tests/`. Databases, supplied Navy files, and generated scratch builds stay out of Git. On Windows with desktop Access, run the standard fresh build and focused tests (also VS Code's default Ctrl+Shift+B task):

```powershell
powershell.exe -NoProfile -File scripts/develop.ps1 -Action BuildAndTest
```

This is the routine development command: it invokes the manifest-driven builder,
creates a new blank database under `build/dev/`, and runs the standard development
checks. Use a copy for manual testing and import existing workspace data through
Settings & Help. Authorized release preparation uses the separate
`scripts/release-check.ps1` gates; building or tagging source does not publish
Blank/Demo files or upload them to Drive.

The `v0.29.0-dev` annotated tag identifies the verified source and records this
local development artifact:

- File: `Swift_NAVFIT_0.29.0-dev_20261022-081737_3ffe14f2.accdb`
- SHA-256: `F8B6883CCFE9DEABF9DE2C811DD19DFEACDD2C6F54B91C5B92B946C5E93994B9`

Generated databases and local verification receipts are not committed. Rebuild
from the tag using the command above; a newly generated Access file can have a
different byte hash.

Start with [current architecture and support boundaries](PROJECT_STATE.md), then [development in VS Code](docs/Development_Workflow.md) for commands and [contributing](CONTRIBUTING.md) for change-specific checks. Release preservation requirements remain in [AGENTS.md](AGENTS.md); the r128 historical-test waiver and unresolved future test policy are recorded in PROJECT_STATE.md. Physical scaling, accessibility, and final printed fit require separate acceptance.

## Guides and project status

- [Swift NAVFIT User Guide - PDF](docs/Swift_NAVFIT_User_Guide.pdf) and [editable online edition](docs/Swift_NAVFIT_User_Guide.md)
- [Architecture description (DoDAF 2.02)](docs/architecture/Swift_NAVFIT_Architecture_Description.md), [DevSecOps alignment](docs/architecture/DevSecOps_Alignment.md), [SBOM](docs/architecture/SBOM.md) and [threat model](docs/architecture/Threat_Model.md)

- [Recent development and verification](planning/RECENT_DEVELOPMENT.md)
- [Current project state](PROJECT_STATE.md), [release history](planning/CURRENT_STATE.md), and [priority backlog](planning/PRIORITY_BACKLOG.md)
- [Summary groups, review routes, and workflow foundations](planning/MAJOR_FOUNDATIONS_GUIDE.md)
- [Workspace guide](planning/UX_WORKSPACE_GUIDE.md)
- [Review assignments and My Work](planning/REVIEW_ROUTING_GUIDE.md)
- [Report tools, unified notes, and verification](planning/REPORT_TOOLS_R123_2026-10-20.md)
- [Historical review-comment service notes](planning/REVIEW_COMMENTS_GUIDE.md)
- [Reporting seniors and averages](planning/REPORTING_SENIOR_GUIDE.md)
- [Workspace transfer and fixed report pages](planning/WORKSPACE_COMPAT_PAGES_R115_2026-10-17.md)
- [Import provenance](planning/PROVENANCE_GUIDE.md)
- [Build and verification history](planning/BUILD_STATUS.md)

This private repository has no selected open-source license. Coordinate use and redistribution with its owner.

Build contributors: see [the managed build pipeline](planning/BUILD_PIPELINE.md) for explicit generator dependencies, input fingerprints and final source/event checks.
