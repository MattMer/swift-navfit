# Swift NAVFIT — current project state

Updated 4 November 2026. This document describes what is true **now**:
architecture, supported behavior, decisions and open gaps. Keep it short.
Per-feature implementation detail and verification receipts up to 0.44.0-dev
are in [planning/PROJECT_STATE_HISTORY.md](planning/PROJECT_STATE_HISTORY.md);
newer verification detail belongs in commit messages, not here.

## Versions

| | Version | Schema | Notes |
| --- | --- | --- | --- |
| Development source | 0.56.0-dev | 48 | Imports schema 35 onward (owner decision, 10 Nov 2026); Settings & Help > About this version and a 90-day Overview reminder, both offline. Field catalogue version 3 (EVAL block 44); restricted-content check (schema 47) with checks that can be turned off (schema 48). [modSwift.bas](src/modSwift.bas) owns `APP_VERSION`/`SCHEMA_VERSION`. |
| Latest verified release | 0.55.0 | 48 | Blank and Demo on the [GitHub release](https://github.com/MattMer/swift-navfit/releases/tag/v0.55.0), 9 Nov 2026. Full release gate passed, including schemas 1–48 and retained populated releases. The last release that imports schemas 1–34. Untested: printed fit, screen readers, an Acrobat-free workstation, a Navy-managed workstation; unsigned. |
| Latest source tag | `v0.55.0` | 48 | The release commit. Earlier: `v0.54.0`, `v0.53.0`, `v0.52.0`, `v0.51.0`, `v0.50.0`, `v0.49.0`, `v0.29.0-dev` (source checkpoint only). r128 (0.27.1-dev, schema 27) remains on the owner's Drive folder. |

## Purpose and boundaries

Windows desktop Microsoft Access application for drafting EVAL, CHIEFEVAL and
FITREP reports, with command member/prior-evaluation history, summary groups and
RSCA, revision history/compare/restore, pending-edit recovery, local review
routes/approvals, report conversations, NAVFIT98A exchange, structured PDF
import and print-ready PDF output.

- Local approval is not an official signature or submission. NAVFIT98A remains
  the destination for native validation and final submission.
- One authoritative workspace file, one contributor at a time. Contributor
  profiles are attribution, not authentication (C03: every launch requires an
  explicit contributor choice).
- **C01 decision (21 Oct):** protection relies on Windows file/folder
  permissions. No database password or encryption. SSNs remain readable in
  storage, revisions, recovery and import provenance; masking is display-only.
- Desktop Access only. Free Access Runtime is out of scope. Tested host:
  Access 16.0.20430.20182 x64.

## Architecture

- **Source:** [src/](src/) holds 112 static VBA modules, JSON catalogues, the
  AutoExec macro, PDF template images and the 70 forms as normalized Access text
  in [src/forms](src/forms). [build-manifest.json](scripts/build-manifest.json)
  lists the modules, forms, generated modules, fingerprinted inputs and six
  resource stages.
- **Build:** [develop.ps1](scripts/develop.ps1) → [build.ps1](scripts/build.ps1)
  creates a fresh database and schema, loads modules, runs the resource stages
  (generated modules, the PDF report, startup properties), loads the forms from
  text, compiles, reopens, and verifies installed source, bindings and view
  contracts. **Forms are source:** edit the text, or change a form in Access on a
  disposable build and re-export it with `scripts/export-form-text.ps1`; see
  [BUILD_PIPELINE.md](planning/BUILD_PIPELINE.md). A build takes about 2 minutes.
- **Build scratch retention:** the standard development and release workflows
  prune generated scratch after 24 hours for successful runs and seven days for
  failed/flaky/unclassified runs. Pins, active workflows, Access sessions and
  supplied input databases are protected; development/release builds, retained
  outputs and workflow receipts remain. See
  [retention commands](docs/Development_Workflow.md#build-scratch-retention).
- **Schema authority:** [major-schema.json](src/major-schema.json) and the base
  builder in `build.ps1` define all DDL. Whenever the schema changes, update both
  together, along with `SCHEMA_VERSION`, metadata and migration.
- **Report fields:** [field-definitions.json](src/field-definitions.json) is the
  versioned authority for storage type, controls, recovery keys, labels,
  validation metadata, NAVFIT/PDF bindings and fixed geometry. Other JSON field
  views are generated from it. [fields.mjs](prototype/fields.mjs) is a retained
  browser reference only.
- **Persistence:**
  - Saved state lives in `ReportHead`, which owns the authoritative
    `ReportRevision`.
  - `RevisionValue` holds the snapshots. `RevisionDate` and `RevisionChoice` hold
    typed dates and choices that keep explicit Missing/Unknown/Conflict states.
  - The `Report` value columns are a checked projection cache. Reads and saves
    stop if the cache disagrees with the revision.
  - Summary-group accounting is relational: `GroupCompletion*`, `RSCountedReport`
    and `GroupLifecycle`.
  - `AuditEvent` indexes the domain histories. `ReportMessage` holds the
    conversations, and legacy note/comment tables stay importable.
  - `WizardPreview` stages recoverable import previews.
  - `MemberCounseling` holds midterm counseling sessions per member;
    `DirectoryMergeCounseling` lets a member merge move and undo them.
  - `ReportSignoff` records each Mark signed and reopening; `DraftDisposal`
    records each draft disposal. A disposed draft keeps its `ReportRevision` row
    (`ValuesDisposed`) so every link to it stays valid.
  - `ReportingSeniorRecord.MemberID` optionally links a senior record to one
    active command member; `DirectoryMergeSenior` lets a member merge move and
    undo that link. A merge is refused when both members back a senior.
- **Report editor:**
  - Uses an unbound buffer with fixed Front/Back NAVPERS geometry.
  - All seven traits stay together on the editing Front Page, and Chief reports
    carry 21 authored performance comments.
  - The left navigation rail is retained.
  - Resizing changes the viewport, not the field layout.
  - Pending values survive page changes, failed saves, VBA reset and restart
    through recovery checkpoints.
  - A line meter under the comments counts lines as NAVFIT98A 2.2.0.332 prints them
    at default pitch (GDI+, its print rectangles and its own limits: comments 18,
    CHIEFEVAL 6; blocks 28/29 3/4; Chief performance comments 2). Longer text gets
    an Information finding that blocks nothing. Unverified against a NAVFIT98A print.
  - Check spelling (button by the comments, or F7 in a narrative block) runs Office
    spelling (acCmdSpelling, local proofing only); accepted corrections go through
    UXSurfaceSetValue as pending edits, and signed reports ask to reopen first.
  - **Mark signed** sets `WorkflowState` to `Signed`, so every save path refuses
    changes. In the editor and group grid the first edit asks whether to reopen
    the report for correction; bulk defaults skip signed reports and list them.
    Completing a summary group asks for one signed date and marks its reports
    signed; leaving Completed reopens them.
  - Report notes are one running message log per report (author and time,
    oldest first) with direct posting and no private drafts. Earlier replies and
    drafts remain stored and importable.
- **Workspace screens** (Reports, Summary groups, My Work, Resources, Settings
  and their siblings) resize through `UXResizeWorkspace`:
  - Narrower than the design, everything scales, and below `UX_COMPACT_WIDTH`
    the rail becomes a drop-down.
  - Wider than the design, only lists, subforms, wide text boxes and wide
    captions stretch, list columns keep their proportions, and content stops at
    16,500 twips. A caption directly above a control with the same left edge
    moves with that control.
  - Widths are capped at Access's 31,680-twip form limit.
  - Pop-up dialogs stay resizable.
- **Entry points:**
  - Startup runs [AutoExec.txt](src/AutoExec.txt) → `SwiftAppStartup` → setup or
    contributor choice → application shell.
  - Every build opens as a plain application window: title "Swift NAVFIT", no
    navigation pane, reduced menus (`AllowFullMenus`, `AllowSpecialKeys`,
    `AllowBuiltInToolbars`, `AllowBreakIntoCode` off), and startup hides the
    ribbon. Right-click menus stay on for Cut/Copy/Paste. This is presentation,
    not protection: holding Shift while opening (`AllowBypassKey` on, by
    decision) gives full Access for development and field investigation.
  - Generated control expressions (`=Fn()`), macros and `Eval` calls are runtime
    callers too, so text search alone cannot prove code unused.
- **Navigation:**
  - Six workspace routes in [modUX.bas](src/modUX.bas) feed the rails.
  - Summary-group Members is the editable grid in `frmUXComparison`.
  - `frmMajorWorkbench` remains a hidden group-context host.
- **PDF output** ([modPDFOutput.bas](src/modPDFOutput.bas)):
  - Fills an Access report over reviewed 300-dpi official template images, using
    Windows GDI font metrics. No Acrobat, PDF.js or Python is used at runtime.
  - Group packets list the summary first, then reports sorted by last name.
  - Overflow goes to labelled additional pages. Findings never block output.
- **Runtime external calls:**
  - MSXML 6, Scripting.FileSystemObject, ADODB.Stream and Windows GDI/crypto
    `Declare`s.
  - No child processes. Backup is a data-only `.accdb` written and verified
    inside Access ([modBackupUI.bas](src/modBackupUI.bas)); restore is the
    Settings & Help import into a new blank workspace.

## Feature status

Release 0.55.0 (9 November 2026, schema 48) adds EVAL block 44
(Qualifications/Achievements) on the Back Page with NAVFIT98A and PDF import,
PDF output and restore from retained NAVFIT98A imports; PDF comments printed at
12 point with a 10-point fallback, and editor boxes that match the printed box;
fit hints in the line meter; the EP and EP+MP spread beside the group RSCA;
drafts and a summary group created together from Reports due; a print preview
from the editor; a restricted-content check with dismissals, allowed phrases and
checks that can be turned off; and screen polish (one font, labelled boxes, one
main action per screen).

Release 0.54.0 adds group directory refresh, editable source-match shading,
report-to-report group navigation, full-NOB mark handling and mixed date/status
counseling entry. Schema 45 adds replaceable working state, retained
narrative/workflow checkpoints and a counselor picker button. It also
adds startup warnings for temporary, synced or read-only files and concurrent
sessions, a clock-safe recovery checkpoint interval, Members-grid checkpoints on
reports under review, and import undo from report History (Operation history is a
read-only log apart from legacy folder undo); clearing block 16 before saving
restores the marks NOB replaced. Block 20 carries a PFA date on CHIEFEVAL only,
entered and shown as MM/YY (stored as the first of the month). Saving a member
profile updates that member's open draft reports. Group defaults follow the
Front Page layout and exclude the reporting senior (from the group's senior
record) and block 29. Previous/Next follow the Members view's sort. One lifecycle function
(`ReportLifecycle`: Archived, Signed, Approved, In review, Draft) backs the group
update and defaults guards, and marking a report signed or archiving it ends its open route or assignment.
Leaving Completed reopens only reports that completion signed. Operation journal
entries move forward only; the retention test fails on an unrecognised revision
pointer. The final schema-45
development build passed canonical BuildAndTest, focused feature checks, the
schema 1–45 import matrix, a populated 0.53.0 upgrade and exact current-schema
retention transfer. These changes are released in 0.54.0 (7 November 2026).
Earlier intermittent native-close and audit/navigation errors passed on isolated
rechecks and the final canonical/focused runs; they remain recorded in the shared
handoff. The migration matrix completed successfully after its monitor timed out.

R15, P01 and M05 are released in 0.53.0 (schema 44). The canonical build and
full release gate passed on the tagged 0.53.0 source.

| ID | Item | Status |
| --- | --- | --- |
| C01 | SSN protection decision | Decided: folder permissions only. |
| C02 | Reviewed duplicate-member merge with guarded undo | Released in r127. |
| C03 | Startup contributor selection | Released in r128. |
| C04 | Effective-dated validation (CH-1 calendars, CH-2 from 26 May 2026, NAVADMIN 132/26 from 3 Jun 2026) | Dev. Dates without coverage produce an explicit Needs-review finding. |
| C05–C07 | Relational RSCA accounting, canonical field catalogue, `ReportHead` authority, typed dates/choices | Dev (schema 28). |
| C08–C10 | Audit index, unified conversations, single navigation list | Dev (schema 29). |
| H01 | Print-ready PDF output, packets, Chief comments and print fidelity | Dev (schema 33). **Not accepted:** physical print fit, Acrobat-absent workstation and the trait-34 grade-4/5 6-point fallback still need checking. |
| H02 | NAVFIT98A export without validation blocking (whole summary group) | Dev (schema 31). |
| M01 | Compiled ACCDE distribution | Unsigned dev candidate only. The compiled runner covers startup, import, saves and PDF output (report PDF and group packet) under trusted automation. Normal-security launch is unresolved and there is no signing certificate. |
| M02 | Recoverable/owned import previews (Resume/Cancel) | Dev (schema 30); follow-ups through 0.44. |
| M03 | Batched reads for folders, history, groups/RSCA, summary list | Dev. |
| M04–M06 | Primary-error diagnostics, required view contracts, `TraitDefinition` table | Dev (schema 31). Page/refresh handlers capture errors before cleanup, retain failure stages and propagate failed Boolean refreshes. Surface refresh/field restoration expose native failures without changing saved or pending values; retry releases the busy state. Postcommit signoff refresh failures report Saved with the committed outcome. Every literal form and control name in VBA, form modules and form properties (bang references, `With` blocks, bound form variables, subforms, `Controls("…")`, probes, open/close targets, sort tags) must resolve against `src/forms`; the static `test-form-references` case in `test-source-contracts` checks this before every build. Code no longer probes for controls a named form always has. Finite computed-name families are expanded from catalogues during source and reopened native checks, with runtime assertions and mutation checks covering navigation, context/detail panels, report fields, traits, shared defaults, roster fields, RSCA rows and sort headings. New dynamic families require a contract. Broader M04 inventory remains in the backlog. |
| R15 | Archived counseling view/restore; PDF reapply after Undo | Released in 0.53.0 (schema 44). Counseling preserves versions/attribution; each PDF attempt retains its own receipt and raw provenance. |
| P01 | Flexible report search, period overlap filters and named filter sets | Released in 0.53.0 (schema 44). Contributor-scoped sets transfer by owner/group GUID; version guards reject stale updates/deletes. |
| M07 | Accessibility (target: Section 508 / WCAG 2.0 AA) | Every input control has an accessible name: its visible label, or a 15-twip `an<Control>` name label that matches the control's visibility. Text contrast meets 4.5:1. test-accessibility guards names, contrast and button captions, with no known gaps. The attended NVDA/JAWS procedure is [docs/Accessibility_Screen_Reader_Test.md](docs/Accessibility_Screen_Reader_Test.md). Under a Windows high-contrast theme, the palette switches to system colors. No screen-reader acceptance. |
| L01–L02 | Descriptive FK names, revision-ID alias queries | Dev (schema 32). |
| Mark signed | Record a signature made outside Swift (per report, per group, or by completing a group); read-only signed reports with a reopen prompt on first edit; optional draft disposal (BUPERSINST 1610.10H Introduction ¶7a) that keeps signed versions and conversations | Dev (schema 35). Final reports are kept: the workspace is the command's reporting senior binder. |
| Reporting seniors | Senior list shows grade, designator, title, UIC and directory link. New reporting senior starts from a member picker (or details by hand); a linked record copies name, grade, designator, UIC and SSN from the member and shows when the directory differs, with Update from directory and Unlink. The title, address and phone stay on the senior record. RSCA history per paygrade lists the entered BOL baseline and each summary group completion (and correction) posted on it, with traits added and the RSCA after; the paygrade table shows each RSCA's source | Dev (schema 36). |
| Counseling | Midterm counseling tracker: chapter 18 schedule by paygrade (current and 2027 tables), records in member history, My Work tab, blocks 30–31 carry-over through Report tools, worksheet disposition after detachment | Dev (schema 34). Printing a pre-filled worksheet is deferred. |

Details and completion criteria: [PRIORITY_BACKLOG.md](planning/PRIORITY_BACKLOG.md).

## Workspace import and compatibility

- Upgrade path: build a new blank, then **Settings & Help → Import data from a
  previous Swift NAVFIT version**. The source is read through read-only DAO and
  its code is never executed. Values are verified before commit, and the import
  is rejected if the target is populated, a source table is unknown, or the source
  schema is newer than the app.
- **The importer accepts source schemas 35–48** (owner decision, 10 Nov 2026).
  Schema 35 is release 0.49.0, the first public release. Older workspaces are
  refused with advice to import them into a 0.55.0 Blank first; 0.55.0 still
  imports schemas 1–48. [modMigration.bas](src/modMigration.bas) hard-codes
  both bounds, and [modWorkspaceImport.bas](src/modWorkspaceImport.bas) checks
  the floor and `SCHEMA_VERSION`. Raise the upper literal whenever the schema
  changes.
- Test evidence (local receipts; see history file):

  | Source schemas | Last verified | Scope |
  | --- | --- | --- |
  | 1–35 | 0.46.0-dev, 25 Oct 2026 | Seeded synthetic source per schema (both schema-27 variants) and populated kept releases for 21–27; production import, independent exact comparison, guards. |
  | 21–29 | 0.29.0-dev, 22 Oct 2026 | Actual retained-schema audit/message matrix. |
  | 28–33 | 0.40.0-dev, 23 Oct 2026 | Actual retained-schema matrix; populated 30–33 major transfers. |

- **Historical-test policy:** every release runs
  [test-workspace-migration.ps1](scripts/test-workspace-migration.ps1) over every
  schema that shipped in a release (35, 36, 41, 44, 45, 48; listed in
  [release-schemas.json](tests/historical-schemas/release-schemas.json)) plus the
  current one, and release-check requires that coverage. Until 0.55.0 it covered
  every schema from 1. Old output files
  are not needed. [tests/historical-schemas](tests/historical-schemas/README.md) holds
  each older schema's DDL:
  - 8–20: exact, from git
  - 21–32: captured from real builds
  - 1–7: reconstructed, because they predate git. These cover tables and the
    documented column changes only; any undocumented column change before
    schema 8 is untested.

  r128's waiver of old-output retesting ended no schema's support.

## Known gaps and test caveats

- `MigrateWorkspace` calls `MigrateLegacyReportNotes` and
  `MigrateReportMessages` through `Application.Eval`. An error raised inside
  them, such as a contributor identity that differs from its audit actor, shows
  an unhandled VBA run-time dialog instead of returning a clean import
  rejection. Fixed in 0.44.0-dev by calling them directly; the import runner's
  audit-identity guard covers it.

- Attended acceptance still outstanding: a NAVFIT98A print check of the narrative
  line meter (backlog H03), physical print fit, an Acrobat-absent
  workstation, normal-security launch and signing, screen reader/keyboard/DPI,
  PDF preview popup visuals.
- No one has tested on a target (Navy-managed) workstation.
- Transient, root cause unknown:
  - one DAO 3059 during fixed-page tests
  - one Access `0xc0000005` crash with a reused fault-test controller
  - an earlier r126 closed-object build failure.
- The C04-era validation-window layout failure (the suite stopped at 98
  assertions) was addressed in 0.37: the client size is now set before asserting,
  and a real 200-twip overlap was fixed. The 0.37 receipt records 141 validation
  assertions passing.

## Superseded approaches (do not reintroduce)

- Release-specific binary assemblers and the duplicate preview builder. Use the
  manifest build and [build-ux-admin-demo.ps1](scripts/build-ux-admin-demo.ps1).
  The retired items are in git history.
- Mandatory conversation resolve/reopen. Retired standalone help/review/checklist
  screens and the Reports Folder-tools screen. Folder data and services remain.
- Treating browser prototype behavior as a production implementation.
