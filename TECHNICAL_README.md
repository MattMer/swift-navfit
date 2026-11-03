# Swift NAVFIT technical, security and system description

This document describes Swift NAVFIT for the people who decide whether it may be
used and who maintain it: command ISSMs and cyber staff, privacy officers,
application approvers, NAVFIT system owners and developers. It covers what the
application is, what data it holds and where, how that data is protected, how
the software is built, tested and released, and what is not yet verified.

For what the application does and how to use it, see the [README](README.md)
and the [user guide](docs/Swift_NAVFIT_User_Guide.md).

**Status.** Swift NAVFIT is a personal, open-source project. It is not an
official U.S. Navy or Department of Defense system, has no authorization to
operate of its own, and is not registered in any DoD or DON inventory. Each
command that uses it decides, through its own ISSM and privacy officer, whether
and where it may run. This document gives them the facts to decide.

**Kept current.** Every change that affects anything described here updates
this document in the same commit (see [Keeping this document
current](#17-keeping-this-document-current)). Sections marked *generated* come
from a fresh build, and a test fails when they, or the data inventory in
section 5.2, no longer match the build.

## 1. Identification

| Item | Value |
| --- | --- |
| Name | Swift NAVFIT |
| Type | Desktop application: one Microsoft Access file per command workspace |
| Purpose | Prepare Navy evaluations, Chief evaluations and fitness reports locally, and track the work around them |
| Owner and maintainer | The project owner, through the public repository (the README names the contact). No government program office. |
| Operated by | The command that adopts it, on its own workstations and storage |
| Classification | Unclassified. Processes personally identifiable information (CUI, category PRVCY). |
| License | Apache License 2.0 |
| Latest release | 0.52.0, schema 41 (3 November 2026), tag `v0.52.0` |
| Security problems | Report privately as described in [SECURITY.md](SECURITY.md) |

Development source (*generated*):

<!-- generated:readme-facts -->
| Fact | Value |
| --- | --- |
| Application version | 0.53.0-dev (`APP_VERSION` in `src/modSwift.bas`) |
| Schema version | 42 |
| Imports accepted | Workspaces of schema 1 to 42, into a target of schema 21 to 42 |
| Tables / enforced relationships | 77 / 168 |
| VBA modules / forms | 114 / 71 |
| Report fields in the catalogue | 100 |
| Test suites by tier | 8 core, 19 extended, 3 helper, 2 manual, 1 release |
| Test cases / VBA test harnesses | 125 / 120 |
<!-- end generated:readme-facts -->

## 2. Purpose and use

Swift NAVFIT helps a command's administrative staff and reporting senior prepare
E1–E6 evaluations (NAVPERS 1616/26), Chief evaluations (1616/27) and officer
fitness reports (1610/2). It:

- edits reports field by field against the official form layouts and the
  BUPERSINST 1610.10H rules in effect for each report's period;
- organizes reports into summary groups and computes trait averages, ITA and
  reporting-senior cumulative average projections;
- runs local review routes and assignments among the command's contributors;
- keeps a member directory, imported from NRRM or MNA rosters, with counseling,
  departures and the reports each member is due;
- exchanges reports with NAVFIT98A and the official PDF forms.

**What it does not do.** Swift does not sign, submit or route reports outside
the command. Official signature and submission stay with the existing process
(NAVFIT98A, the official forms and BUPERS Online). Swift has no network
connection, authenticates no one, and stores no credentials.

**Users.** Command administrative staff, the reporting senior, and local
reviewers (for example, division officers and department heads). Rated members
never use it. One workspace file serves one command, used by one person at a
time.

## 3. Registration and approval considerations

These are the facts approvers usually ask for, with the maintainer's
assessment. The decisions belong to the command and its approvers.

| Question | Facts |
| --- | --- |
| System or application? | An application. It runs inside the workstation's existing environment (for example, NVD) with no servers, services, network connections or interfaces to other systems. DoD inventory guidance (DITPR) generally does not register applications or common user tools that lack "the comprehensive integrated characteristics of an IT system". |
| Defense business function? | Swift supports a human-resources activity (preparing performance reports) but performs no official transaction: reports are signed and submitted through the existing process. Whether this makes it subject to registration as a defense business system (10 U.S.C. § 2222) is for the command's approvers to decide. |
| Software approval | Navy-managed networks normally require software to be approved for use (for example, through DADMS / DITPR-DON). Swift runs inside Microsoft Access, which is already approved; the command decides whether a custom Access database needs its own entry. |
| Authorization (ATO) | None of its own. It relies on the authorization of the workstation and network it runs on. |
| National Security System questions | No intelligence, cryptologic, command-and-control or weapons function. No classified information. |
| Mission criticality | Mission support. If Swift is unavailable, reports can still be prepared in NAVFIT98A. |
| Interfaces to other systems | None. Data moves only through files the user chooses (section 4.3). |
| Number of users | One workspace per command; one user at a time. |
| Section 508 | Automated accessibility checks pass; attended screen-reader acceptance is not yet done (section 13). |
| Contingency | Verified backups and restore (section 8). NAVFIT98A remains the fallback. |
| Code signing | Not yet signed; a DoD PKI code-signing certificate is needed (section 6.7). |

## 4. Operating environment and architecture

### 4.1 Platform

| | |
| --- | --- |
| Operating system | Windows |
| Required software | Full desktop Microsoft Access 2016 or later, including Microsoft 365, 64-bit tested. The free Access Runtime is not supported. |
| Tested host | Access 16.0.20430.20182 x64 |
| Language | VBA, with Access forms, queries and one macro (`AutoExec`) |
| Installation | None. Copy the `.accdb` to approved storage and open it. No administrator rights, registry changes, services or drivers. |
| Network | None. No telemetry, update check, sync or web service. |
| Third-party runtime code | None. Everything called at run time is part of Windows or Microsoft Office ([SBOM](docs/architecture/SBOM.md)). |

### 4.2 Architecture

Swift NAVFIT is a single-file desktop Access application: the user interface,
code and data all live in one workspace `.accdb`. There is no separate front
end and back end.

- **Source is text.** [src/](src/) holds the VBA modules, the forms as
  normalized Access text in [src/forms](src/forms), the `AutoExec` macro, JSON
  catalogues and the PDF template images. No `.accdb` is committed.
- **Builds are fresh and reproducible from source.**
  [develop.ps1](scripts/develop.ps1) calls [build.ps1](scripts/build.ps1), which
  creates an empty database, creates the schema, loads the modules, runs the
  generated-resource stages listed in
  [build-manifest.json](scripts/build-manifest.json), loads the forms, compiles,
  reopens, and verifies the installed source, control bindings and view
  contracts. A build takes about 2 minutes and writes a `.build.json` receipt
  that fingerprints every input.
- **The schema has one authority.** [major-schema.json](src/major-schema.json)
  and the base builder define all tables and relationships; every schema change
  also bumps `SCHEMA_VERSION` and the importer's accepted range.
- **Report fields have one authority.**
  [field-definitions.json](src/field-definitions.json) defines each field's
  storage type, control, labels, limits, validation metadata, NAVFIT98A and PDF
  bindings and fixed page geometry. Other field views are generated from it.
- **Persistence is revision-based.** Each save writes a new `ReportRevision`
  with a full snapshot in `RevisionValue`, typed dates and choices in
  `RevisionDate` and `RevisionChoice`, and a field-level change log.
  `ReportHead` points at the authoritative revision. The `Report` value columns
  are a projection cache; reads and saves stop if the cache disagrees with the
  revision.
- **Summary-group accounting is relational:** completions, counted reports and
  lifecycle states are stored as rows with foreign keys, so averages can be
  reconciled by query.
- **Audit and history:** `AuditEvent` indexes the domain histories (revisions,
  review decisions, merges, sign-offs, draft disposals). Old audit, note and
  comment records are kept and stay importable.

The [architecture description](docs/architecture/Swift_NAVFIT_Architecture_Description.md)
and [PROJECT_STATE.md](PROJECT_STATE.md) give the full picture.

### 4.3 Data exchange (file interfaces)

Every exchange is a file the user chooses. None is automatic.

| Interface | Direction | Format | Safeguards |
| --- | --- | --- | --- |
| NAVFIT98A exchange | in and out | Access database | Read-only DAO for imports; field coverage checked; the source is never modified |
| Official PDF output | out | PDF on the NAVPERS fillable templates | File names never contain SSNs |
| Structured PDF import | in | PDF with embedded form data | Preview before anything is applied; provenance kept |
| NRRM roster import | in | CSV | Match preview; manual changes on profiles are protected |
| MNA roster import | in | Excel workbook (`.xlsx`), read as text by the Access engine | Headers checked; preview before anything is applied |
| Earlier-workspace import | in | Access database (every earlier schema) | Read-only; its code is never run; every value verified before commit; blank target only |
| Backup and restore | out and in | Access database with a verification record | Row-by-row verification; restore only into a blank workspace |

The [SV-1 view](docs/architecture/views/SV-1.md) lists the test runner for each
interface.

### 4.4 Repository layout

| Path | Contents |
| --- | --- |
| `src/` | Application source: VBA modules, forms as text, JSON catalogues, the `AutoExec` macro, PDF template images. |
| `scripts/` | Build, verification, release and test-suite scripts (PowerShell, a few Python helpers). |
| `scripts/cases/` | One script per test case; suites run them in their own processes. |
| `tests/` | VBA test harnesses, the test-suite registry (`runners.json`) and historical schema fixtures. |
| `docs/` | User guide, task guides, accessibility procedure, architecture documents. |
| `planning/` | Backlog, design notes, findings and verification history. |
| `prototype/` | A retained browser reference for field definitions. Not production code. |

Databases, generated builds, supplied Navy files and reference forms are kept
out of Git.

## 5. Data and privacy

### 5.1 Personal data held

| Category | Elements | Source |
| --- | --- | --- |
| Rated member identity | Name, report name, rank or rate, designator, SSN, DoD ID (from MNA rosters and report fields), UIC, ship or station | Typed, NAVFIT98A, PDF, NRRM or MNA rosters |
| Assignment | Billet, PRD, NEC/NOBC, reserve unit data, date reported, departure dates and reasons | NRRM or MNA rosters; typed |
| Contact details | Address, phone and e-mail where the NRRM roster includes them | NRRM roster |
| Evaluation content | Trait grades, promotion recommendation, comments, career recommendations, qualifications, counseling dates and counselor, prior evaluation history | Typed, NAVFIT98A, PDF |
| Reporting senior and raters | Name, grade, designator, title, UIC, SSN, DoD ID, address, phone, e-mail; rater, senior rater and counselor names | Typed, directory, PDF |
| Review work | Notes, messages, review decisions and their free text | Typed by contributors |
| Contributors | Display names (optionally linked to a directory member) | Typed or from the directory |

Swift creates no new personal data beyond what the evaluation process already
uses. It does not collect data from rated members directly.

### 5.2 Data inventory by table

Every table in the workspace, by what it holds. A test fails when a table is
added or removed without updating this list.

<!-- data-inventory -->
| Area | Tables | Personal data |
| --- | --- | --- |
| Member directory and rosters | `CommandMember`, `MemberReport`, `PriorEvaluation`, `PriorEvaluationTrait`, `MemberCounseling`, `MemberDeparture`, `RosterImport`, `RosterPresence`, `ReportDueDismissal` | Yes: identity, SSN, roster data (including DoD ID and contact details), prior grades, counseling, departures |
| Reporting seniors | `ReportingSeniorRecord`, `ReportingSeniorValue`, `RSCABaseline`, `GroupRSCABaseline`, `RSCountedReport` | Yes: senior identity, SSN, contact details (including e-mail); cumulative averages |
| Directory merge history | `DirectoryMerge`, `DirectoryMergeReport`, `DirectoryMergePrior`, `DirectoryMergeCounseling`, `DirectoryMergeSenior` | Links between member records before and after a merge |
| Reports and revisions | `Report`, `ReportHead`, `ReportRevision`, `RevisionValue`, `RevisionChange`, `RevisionDate`, `RevisionChoice`, `ReportEvent`, `ReportSignoff`, `DraftDisposal`, `RecoveryDraft` | Yes: every report field in every saved version, unsaved checkpoints, sign-off and disposal records |
| Imports | `ImportReceipt`, `ImportRawValue`, `WizardPreview` | Yes: raw values as imported (including SSNs) and import provenance |
| Bulk changes and defaults | `ChangeBatch`, `ChangeItem`, `FieldDefault`, `FieldDefaultLink`, `MeetingProposal` | Yes: field values before and after a change, shared default values |
| Validation | `ValidationRun`, `ValidationFinding`, `RulePack`, `RuleDefinition`, `TraitDefinition` | Finding messages can quote report values; rule tables hold policy only |
| Summary groups | `SummaryGroup`, `GroupMember`, `GroupSnapshot`, `GroupSnapshotMember`, `GroupCompletion`, `GroupCompletionReport`, `GroupCompletionCurrent`, `GroupLifecycle`, `ComparisonSet`, `ComparisonMember` | Membership and accounting links to reports |
| Review and collaboration | `ReviewTask`, `ReviewTaskEvent`, `ReviewComment`, `ReviewCommentEvent`, `ReportNote`, `ReportNoteDraft`, `ReportMessage`, `RouteTemplate`, `RouteTemplateStep`, `RouteRun`, `RouteRunStep`, `RouteRunEvent` | Free text written by contributors, which may mention members |
| Contributors and audit | `UserProfile`, `AuditActor`, `AuditEvent`, `OperationJournal` | Contributor names; operation descriptions (which can name members and show the last four SSN digits) |
| Summary letters | `ISIC` | ISIC names, UICs and mailing addresses (organizations, not people) |
| Workspace and library | `Workspace`, `AppMetadata`, `Folder`, `FolderMove`, `FolderBatchChange`, `LibraryPreference`, `NarrativeSnippet` | Command name; reusable comment text written by users |
<!-- end data-inventory -->

### 5.3 Where data goes

- **Inside the workspace file.** All of section 5.2.
- **Files the user creates:** report PDFs, NAVFIT98A exports, verified backups.
  They hold the same personal data as the workspace and must be kept on storage
  approved for it. Swift writes them only to the folder the user chooses.
- **Temporary working copies.** Swift writes short-lived files to the user's
  Windows temporary folder (`%TEMP%`) and deletes them when the operation
  finishes: a copy of an older NAVFIT98A database being upgraded for import, the
  NAVFIT98A export template while an export is built, and a cache for the trait
  comparison grid. The first and last can contain personal data. A crash can
  leave one behind; the user profile's temporary folder must therefore sit on
  storage approved for the same data.
- **Nowhere else.** No network transmission and no logging to other locations.

### 5.4 Retention and disposal

Swift deletes nothing on its own schedule. Reports are archived, not deleted,
and every saved version is kept, because the workspace is the reporting
senior's long-term binder. A draft that was never needed can be disposed of; the
disposal is recorded. Disposing of the workspace file, its backups and exports
follows the command's records and PII disposal procedures.

### 5.5 Items for the command's privacy officer

- Whether the command's use is covered by its existing privacy documentation
  for performance evaluations (privacy impact assessment and system of records
  notice), or needs its own.
- Where the workspace, backups and exports may be stored, and who may have
  access (Swift's access control is the folder's permissions; section 6.1).
- CUI marking and handling for printed and exported reports.
- Whether data-at-rest encryption of the storage location is required and
  provided (Swift adds none of its own; section 6.6).

## 6. Security controls

### 6.1 Access control

Swift relies on the Windows permissions of the folder that holds the workspace.
Anyone who can open the file can read and change everything in it. There is no
database password or encryption: this is a documented decision (C01 in the
[backlog](planning/PRIORITY_BACKLOG.md)) so that the command's existing
file-share controls are the single access boundary.

### 6.2 Identification and attribution

Contributor profiles record who made each change. They are attribution, not
authentication: anyone who can open the file can choose any contributor. The
active contributor cannot be deleted. Deleted contributors keep their name on
their past work, cannot be selected, and are removed from review routes.

### 6.3 Audit and accountability

Every save is a revision recording the contributor and time, with a field-level
change log. Review decisions, route events, messages, sign-offs, draft
disposals, merges, imports and bulk changes are recorded. `AuditEvent` indexes
these histories. History is never rewritten; restoring an older version writes
a new revision.

### 6.4 Integrity

- Saves run in a transaction; a failed save rolls back and keeps the pending
  values for retry.
- Table validation rules and relationships enforce structure even outside
  Swift.
- Signed reports refuse every save path until reopened.
- Direct table edits in Access are not cryptographically detectable (threat
  model T1); an optional integrity hash is a recommendation.

### 6.5 Untrusted input

Imports (NAVFIT98A databases, earlier Swift workspaces, PDFs, NRRM CSV, MNA
workbooks) are read as data:

- earlier workspaces and NAVFIT98A databases are opened read-only and their code
  is never run;
- every XML parser disables DTDs and external resolution, and a source check
  fails any parser that doesn't;
- the PDF reader limits file size, decoded size, nesting and object counts;
- `test-malformed-imports` and `test-external-import` run malformed and hostile
  inputs against every importer;
- imports preview before applying, and a failed import rolls back.

### 6.6 Confidentiality of stored data

- SSNs are stored readable in current values, revisions, recovery checkpoints
  and import provenance. Masking in lists and previews is display only.
- SSNs are kept out of generated file names.
- Swift adds no encryption at rest; it relies on the protection of the
  workstation and storage it runs on.

### 6.7 Software integrity and supply chain

- **No third-party runtime code.** The [SBOM](docs/architecture/SBOM.md)
  (also [CycloneDX](docs/architecture/sbom.cdx.json)) is generated from each
  build and lists every Windows API and COM class the code calls.
- **No network use.** Runtime external calls are limited to MSXML 6,
  `Scripting.FileSystemObject`, `ADODB.Stream`, `Shell.Application` (to open a
  generated PDF) and Windows GDI and crypto API declarations. Swift starts no
  other programs.
- **Releases** are published on GitHub with SHA-256 hashes. Builds are made
  fresh from tracked source and checked against it before release.
- **Code signing is not done yet.** Users must choose **Enable Content**, and a
  Navy-managed workstation's Office policy may block that. The
  [code-signing plan](docs/architecture/Threat_Model.md#5-code-signing-plan-m01)
  needs a DoD PKI code-signing certificate requested through the command ISSM.

### 6.8 Configuration

Builds open as a plain application window with the navigation pane, full menus
and ribbon hidden. This is presentation, not protection: holding Shift while
opening bypasses it (`AllowBypassKey` stays on, by decision, for support).

Swift changes one Office setting, and only briefly: during its own spell check
it turns off Access's "Ignore words in UPPERCASE" so all-capital Navy terms are
checked, then restores the user's value.

### 6.9 Error handling

Error messages name the operation and stage; most do not include report values,
but some validation messages quote the value they reject (threat model I4,
partly mitigated). Problems are reproduced with synthetic data; the project
never accepts real personnel data in bug reports ([SECURITY.md](SECURITY.md)).

## 7. Threats and residual risk

The [threat model](docs/architecture/Threat_Model.md) applies STRIDE to Swift's
assets, trust boundaries and interfaces. The main residual risks:

| Risk | Status | Owner |
| --- | --- | --- |
| Someone with folder access reads or changes the workspace, or picks another contributor's name | Accepted by design; controlled by folder permissions | Command |
| The workspace or a backup is copied off approved storage | Command storage and handling controls | Command |
| A modified copy of the release file circulates | Partly mitigated by published hashes; code signing planned | Maintainer and ISSM |
| Direct table edits bypass Swift's rules | Partly mitigated by table rules and relationships; integrity hash optional | Maintainer |
| Enabling content runs Swift's code with the user's rights | Partly mitigated: no network, no admin rights; signing would narrow trust to one publisher | Maintainer and ISSM |

Open recommendations: sign releases; optionally hash each saved revision.

## 8. Data integrity and recovery

- Unsaved edits are checkpointed separately from saved revisions, so they
  survive page changes, failed saves, a VBA reset and a restart.
- Restoring an earlier version writes a new revision; history is never
  rewritten.
- Draft disposal keeps the revision rows so every link stays valid.
- Backups are data-only `.accdb` copies written and verified inside Access.
  Restore is an import into a new blank workspace.

## 9. Compatibility and upgrades

- The importer accepts source workspaces of every earlier schema, from 1 to the
  current one, into a target of schema 21 or later. It refuses a populated
  target, an unknown source table, or a source newer than the application.
- Every value is compared with the source before the import commits.
- Every release runs the workspace-migration suite over every schema.
  [tests/historical-schemas](tests/historical-schemas/README.md) holds each
  older schema's DDL: 8–20 exact from Git, 21 onward captured from real builds,
  1–7 reconstructed (they predate Git; undocumented column changes there are
  untested).

## 10. Testing and verification

- **Static checks before every build:** VBA source rules
  (`check-vba-source.ps1`), the build plan and test registry, canonical field
  definitions, and form and control references (every literal form and control
  name in the code must exist, with mutation checks proving a rename or removal
  fails).
- **Native tests** drive Access through COM against a fresh build. Each case
  runs on its own copy of the build, and any failure fails its suite. Core
  tests run with every build (`develop.ps1 -Action BuildAndTest`, about 5
  minutes); the extended tier runs before a release or a risky change
  (`-Action ExtendedTest`, about an hour); the release tier runs on release
  candidates only.
- **Coverage includes** persistence and rollback, recovery after a simulated
  crash, imports of every historical schema, NAVFIT98A round trips, malformed
  imports, PDF output fixtures, validation rule packs, review routes, responsive
  layout, accessibility names and contrast, and the architecture views.
- **Generated documents are checked.** Architecture views, the SBOM, state
  lists and the generated parts of this document are regenerated from a build,
  and a test fails if they drift from it.
- **Receipts.** Each run writes a `verification.json` receipt under
  `build/workflow` with the build's hash. Receipts are local and not committed;
  verification results are recorded in commit messages.

## 11. Release process

1. Build a new Blank from the tagged source, and a synthetic Demo with
   [build-ux-admin-demo.ps1](scripts/build-ux-admin-demo.ps1).
2. Run [release-check.ps1](scripts/release-check.ps1) on both files. It checks
   they match the current source, requires the Blank to be empty, and runs the
   release gates: member merge and history, every historical schema import,
   populated transfers with rejection and rollback, fixed report pages, workflow
   and recovery, engine exactness and message transfer.
3. Publish both files on a GitHub release with their SHA-256 hashes, tagged
   `vX.Y.Z`. Releases tagged `-dev` are unverified test builds.

A rebuilt `.accdb` can have a different byte hash from the published one; the
published hashes identify the released files. The Demo contains only synthetic
data and is the right file to share for evaluation.

## 12. Performance and limits

Measured on a synthetic workspace of 4,700 reports:

| Operation | Time |
| --- | --- |
| Save a report | about 43 ms |
| Create a report | about 70 ms |
| Load the Reports list | about 0.4 s |
| Load a 250-member summary group | about 1.2 s |
| Open the report editor | about 1.1 s |
| PDF packet for a 250-member group | about 46 s |

Access limits a database to 2 GB. That workspace was 272 MB after compact and
repair, about 26 KB per save, which puts the limit at roughly 78,000 saves
(about 10 years for a command writing 1,300 reports a year). Settings & Help
shows the file size and warns from 1.5 GB.

## 13. Accessibility (Section 508)

The target is the DoD Section 508 minimum (36 CFR 1194, WCAG 2.0 AA).
Automated checks pass: contrast pairs at 4.5:1 or better, every input and
button named, keyboard order complete, Windows contrast themes followed. Not yet
done: attended screen-reader testing (NVDA, JAWS), display scaling, and a formal
conformance report ([screen-reader test procedure](docs/Accessibility_Screen_Reader_Test.md)).
Some actions are pointer-only, for example double-clicking a report people
field to fill it from the directory; typing the value is the keyboard
alternative.

## 14. Known limitations and open items

- **Attended acceptance not yet done:**
  - printed fit against the official forms;
  - a NAVFIT98A print check of the narrative line meter;
  - screen readers (NVDA and JAWS), keyboard, and display scaling;
  - a workstation without Adobe Acrobat;
  - a Navy-managed workstation, normal-security launch and signing.
- **Single user.** One workspace file per command, one person at a time.
  Copies do not sync; shared storage (SharePoint Lists) is proposed, not built.
- **Validation is bounded** to the implemented rule packs. Dates outside their
  coverage produce a Needs-review finding rather than a verdict. NAVFIT98A
  validation is still required after export.
- **Roster formats.** NRRM and MNA are supported. MNA lists no officer rank, so
  new officer profiles from MNA start without one.
- **Intermittent test failures, root cause unknown** (backlog T05): a handful of
  tests have failed once and passed on retry. Diagnostics are captured the next
  time one recurs.

The full list is in [PROJECT_STATE.md](PROJECT_STATE.md#known-gaps-and-test-caveats)
and the [priority backlog](planning/PRIORITY_BACKLOG.md).

## 15. Architecture and compliance documents

- [Architecture description](docs/architecture/Swift_NAVFIT_Architecture_Description.md),
  aligned with DoDAF 2.02 and structured per ISO/IEC/IEEE 42010, with 14
  fit-for-purpose views in [docs/architecture/views](docs/architecture/views).
  Five views are generated from the source and a build. It is DoDAF-aligned,
  not certified conformant: formal conformance needs DoD architecture tooling
  (DM2/PES).
- [DevSecOps alignment matrix](docs/architecture/DevSecOps_Alignment.md): every
  activity in the DoD Enterprise DevSecOps Guidebook, marked done (with
  evidence), partial, command-owned or not applicable.
- [Software bill of materials](docs/architecture/SBOM.md), also as
  [CycloneDX 1.5](docs/architecture/sbom.cdx.json), generated from the build.
- [Threat model](docs/architecture/Threat_Model.md): STRIDE over the assets,
  trust boundaries and file interfaces, with mitigations, recommendations and
  the code-signing plan.
- [Screen-reader test procedure](docs/Accessibility_Screen_Reader_Test.md) for
  Section 508 acceptance.
- NAVFIT98A evidence: [package findings](planning/NAVFIT_PACKAGE_FINDINGS.md),
  [validation comparison](planning/NAVFIT_VALIDATION_COMPARISON_2026-10-14.md)
  and [print-engine findings](planning/NAVFIT_PRINT_ENGINE_FINDINGS.md).

## 16. Building from source

You need Windows, full desktop Access, and Windows PowerShell 5.1. Run from the
repository root:

```powershell
powershell.exe -NoProfile -File scripts/develop.ps1 -Action BuildAndTest
```

This builds a new blank database under `build/dev/` and runs the core tests. Use
`-Action ExtendedTest` for the extended tier. To try the build with data, create
a Demo from it:

```powershell
powershell.exe -NoProfile -File scripts/build-ux-admin-demo.ps1 -TemplatePath build/dev/<build>.accdb -OutputPath build/dev/<demo>.accdb
```

Contribution rules, change-specific checks and Access automation safety are in
[CONTRIBUTING.md](CONTRIBUTING.md). Commands and test suites are described in
[Development in VS Code](docs/Development_Workflow.md), and the build pipeline
in [BUILD_PIPELINE.md](planning/BUILD_PIPELINE.md).

## 17. Keeping this document current

Update this document in the same commit as any change to:

- the data Swift holds, imports or exports (sections 5 and 4.3);
- security behaviour: access, attribution, audit, integrity, input handling,
  encryption, signing or dependencies (section 6);
- the platform, installation or interfaces (sections 3 and 4);
- testing, release or known limitations (sections 10, 11 and 14);
- a release (section 1).

Two checks enforce part of this. `scripts/generate-architecture-views.ps1`
writes the generated facts in section 1 from a fresh build, and
`test-architecture-views` fails when they are stale or when the data inventory
in section 5.2 does not list exactly the build's tables. Add a line to the
document history below for each substantive change.

## 18. Document history

| Date | Change |
| --- | --- |
| 3 Nov 2026 | Schema 42: ISICs and the reporting senior's e-mail, printed on summary letters (sections 5.1 and 5.2). |
| 3 Nov 2026 | Release 0.52.0. |
| 3 Nov 2026 | Spell check briefly changes Access's uppercase setting (section 6.8). |
| 2 Nov 2026 | Expanded from the technical README into a security and system description: identification, registration and approval considerations, data inventory and privacy, security controls, residual risk, accessibility, and the keep-current rule. Covers schema 41 (roster imports from NRRM and MNA, departures, reports due, member date reported, contributor deletion and directory pickers). |

## Further reading

- [Current project state](PROJECT_STATE.md) and its [history](planning/PROJECT_STATE_HISTORY.md)
- [Priority backlog](planning/PRIORITY_BACKLOG.md), [recent development](planning/RECENT_DEVELOPMENT.md) and [release history](planning/CURRENT_STATE.md)
- [Build and verification history](planning/BUILD_STATUS.md)
- [Summary groups, review routes and workflow foundations](planning/MAJOR_FOUNDATIONS_GUIDE.md)
- [Workspace guide](planning/UX_WORKSPACE_GUIDE.md)
- [Review assignments and My Work](planning/REVIEW_ROUTING_GUIDE.md)
- [Reporting seniors and averages](planning/REPORTING_SENIOR_GUIDE.md)
- [Import provenance](planning/PROVENANCE_GUIDE.md)

## License

Apache License 2.0; see [LICENSE](LICENSE) and [NOTICE](NOTICE). The SBOM records
the license in its CycloneDX metadata. Swift NAVFIT is a personal project, built
on its author's own time and equipment, and is not an official U.S. Navy or
Department of Defense system.
