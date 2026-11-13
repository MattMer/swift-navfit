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
| Latest release | 0.55.0, schema 48 (9 November 2026), tag `v0.55.0` |
| Security problems | Report privately as described in [SECURITY.md](SECURITY.md) |

Development source (*generated*):

<!-- generated:readme-facts -->
| Fact | Value |
| --- | --- |
| Application version | 0.56.0-dev (`APP_VERSION` in `src/modSwift.bas`) |
| Schema version | 48 |
| Imports accepted | Workspaces of schema 35 to 48, into a target of schema 21 to 48 |
| Tables / enforced relationships | 83 / 169 |
| VBA modules / forms | 122 / 74 |
| Report fields in the catalogue | 101 |
| Test suites by tier | 8 core, 19 extended, 3 helper, 2 manual, 1 release |
| Test cases / VBA test harnesses | 126 / 123 |
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
  defines all tables and relationships; every schema change
  also bumps `SCHEMA_VERSION` and the importer's accepted range.
- **Report fields have one authority.**
  [field-definitions.json](src/field-definitions.json) defines each field's
  storage type, control, labels, limits, validation metadata, NAVFIT98A and PDF
  bindings and fixed page geometry. Other field views are generated from it.
- **Persistence is revision-based.** Each save writes a new `ReportRevision`
  with a full snapshot in `RevisionValue`, typed dates and choices in
  `RevisionDate` and `RevisionChoice`, and a field-level change log.
  Routine working states replace their unreferenced predecessor; narrative
  saves and workflow checkpoints are retained. Unique head identities still
  reject stale writes even when the preceding working snapshot is removed.
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
| Earlier-workspace import | in | Access database (schema 35 onward) | Read-only; its code is never run; every value verified before commit; blank target only |
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
| Directory merge history | `DirectoryMerge`, `DirectoryMergeReport`, `DirectoryMergePrior`, `DirectoryMergeCounseling`, `DirectoryMergeSenior`, `DirectoryMergeLink` | Links between member records before and after a merge |
| Reports and revisions | `Report`, `ReportHead`, `ReportRevision`, `RevisionValue`, `RevisionChange`, `RevisionDate`, `RevisionChoice`, `RevisionRetention`, `ReportEvent`, `ReportSignoff`, `DraftDisposal`, `RecoveryDraft` | Yes: current report fields, retained versions, pending recovery checkpoints, sign-off and disposal records |
| Imports | `ImportReceipt`, `ImportRawValue`, `WizardPreview` | Yes: raw values as imported (including SSNs) and import provenance |
| Bulk changes and defaults | `ChangeBatch`, `ChangeItem`, `FieldDefault`, `FieldDefaultLink`, `MeetingProposal` | Yes: field values before and after a change, shared default values |
| Validation | `ValidationRun`, `ValidationFinding`, `RulePack`, `RuleDefinition`, `TraitDefinition` | Finding messages can quote report values; rule tables hold policy only |
| Content-check decisions | `RestrictedDismissal`, `RestrictedException`, `RestrictedCheckOff` | Yes: dismissed findings keep the matched words from a report (which can name family members or medical terms) with the report, contributor and reason; allowed phrases are command wording; checks turned off are rule IDs only |
| Summary groups | `SummaryGroup`, `GroupMember`, `GroupSnapshot`, `GroupSnapshotMember`, `GroupCompletion`, `GroupCompletionReport`, `GroupCompletionCurrent`, `GroupLifecycle`, `ComparisonSet`, `ComparisonMember` | Membership and accounting links to reports |
| Review and collaboration | `ReviewTask`, `ReviewTaskEvent`, `ReviewComment`, `ReviewCommentEvent`, `ReportNote`, `ReportNoteDraft`, `ReportMessage`, `RouteTemplate`, `RouteTemplateStep`, `RouteRun`, `RouteRunStep`, `RouteRunEvent` | Free text written by contributors, which may mention members |
| Contributors and audit | `UserProfile`, `AuditActor`, `AuditEvent`, `OperationJournal` | Contributor names; operation descriptions (which can name members and show the last four SSN digits) |
| Summary letters | `ISIC` | ISIC names, UICs and mailing addresses (organizations, not people) |
| Saved report filters | `SavedReportFilter` | Contributor GUID, named search criteria, period bounds and group scope; search text and filter names may contain member names |

| Workspace and library | `Workspace`, `AppMetadata`, `Folder`, `FolderMove`, `FolderBatchChange`, `LibraryPreference`, `NarrativeSnippet` | Command name; reusable comment text written by users |
<!-- end data-inventory -->

Saved report-filter sets are scoped by contributor GUID for convenience, not
authentication. They transfer as workspace data, including any names in search
text. Schema 44 also adds a nullable `ImportReceipt.ReceiptAttempt`: legacy
receipts retain an unspecified attempt, and new import attempts have distinct
sequence numbers. The source/record/attempt index retains each receipt while
the PDF adapter refuses a second active import; Undo permits a fresh reviewed
attempt without rewriting the earlier receipt's source identity or provenance.

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

Reports are archived, not deleted. A successful save replaces the superseded,
unreferenced working snapshot created under the schema-45 policy. Narrative
versions and workflow checkpoints are retained, along with existing/imported
history and snapshots needed by review, undo, signing, messages or previews.
After Mark signed, confirmed draft disposal can remove eligible draft contents;
signed versions, conversations, group/accounting history and the disposal record
remain. Disposing of the workspace file, its backups and exports follows the
command's records and PII disposal procedures.

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

Startup warns, without blocking, when the workspace is open from a temporary
folder, sits in a sync-service folder (by path or cloud-file attribute), is
read-only, or is already open in another session (from the Access lock file's
connected-user roster, which gives computer names only). A session whose
contributor another session changed is refused with that explanation; nothing
is attributed to the other session's contributor.

### 6.3 Audit and accountability

Changed saves record the current contributor and time. Retained narrative and
checkpoint versions include field-level differences; routine working snapshots
and their temporary change rows are replaced. Review decisions, route events,
messages, sign-offs, draft disposals, merges, imports and bulk changes are
recorded. `AuditEvent` indexes retained revision and domain histories. Restoring
an older version saves its values as current without rewriting retained history.
Members-grid saves on a report under review, and the states before and after an
import into an existing draft, are retained checkpoints. Restoring the state before
an import, while the import is the latest change, also marks its receipt undone.

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
  `Scripting.FileSystemObject`, `ADODB.Stream` and Windows GDI and crypto API
  declarations. Swift starts no other programs itself: it opens a generated PDF
  or its folder through Access's own `Application.FollowHyperlink`, and always
  shows the folder path to copy. The source check rejects shell objects
  (`Shell.Application`, `WScript.Shell`, `Shell`, `ShellExecute`), which managed
  workstations flag as macro threats.
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
| Two sessions open one file, or the file sits in a temporary or synced folder | Partly mitigated: startup warns and a changed contributor is refused; one person at a time remains a procedure | Command |
| The workspace or a backup is copied off approved storage | Command storage and handling controls | Command |
| A modified copy of the release file circulates | Partly mitigated by published hashes; code signing planned | Maintainer and ISSM |
| Direct table edits bypass Swift's rules | Partly mitigated by table rules and relationships; integrity hash optional | Maintainer |
| Enabling content runs Swift's code with the user's rights | Partly mitigated: no network, no admin rights; signing would narrow trust to one publisher | Maintainer and ISSM |

Open recommendations: sign releases; optionally hash each saved revision.

## 8. Data integrity and recovery

- Unsaved edits are checkpointed separately from saved revisions, so they
  survive page changes, failed saves, a VBA reset and a restart. The 30-second
  checkpoint interval restarts if the workstation clock goes back.
- Restoring an earlier version writes a new revision; history is never
  rewritten.
- Draft disposal keeps the revision rows so every link stays valid.
- Backups are data-only `.accdb` copies written and verified inside Access.
  Restore is an import into a new blank workspace.

## 9. Compatibility and upgrades

- The importer accepts source workspaces from schema 35 (release 0.49.0) to the
  current one, into a target of schema 21 or later. Older workspaces are
  refused with advice to import them into a 0.55.0 Blank first, which still
  imports schemas 1–48. It refuses a populated
  target, an unknown source table, or a source newer than the application.
- Referenced ISIC organizations are copied before summary groups so their links
  remain valid during schema 42 and later imports.
- Every value is compared with the source before the import commits.
- Every release runs the workspace-migration suite over every schema that
  shipped in a release, listed in
  [release-schemas.json](tests/historical-schemas/release-schemas.json), plus the
  current one.
  [tests/historical-schemas](tests/historical-schemas/README.md) holds each
  older schema's DDL: 8–20 exact from Git, 21 onward captured from real builds,
  1–7 reconstructed (they predate Git; undocumented column changes there are
  untested).

A group directory refresh uses only confirmed member links and the selected
reporting senior record. It confirms replacements, checks source/report/group
versions and preview ownership, preserves unentered member values, applies
cleared senior fields, skips signed reports and applies eligible changes atomically.
It rechecks source versions before commit and only pulls values into reports;
it does not enrich member profiles from older report values during this action. Gray editor fields indicate a current value match
with a linked record; they remain editable and do not establish import provenance
or prevent local overrides. Previous/Next uses an explicit summary group and the
normal pending/save/cancel guards. Block 30 stores either an ISO date or the
existing NOT REQ / NOT PERF statuses; its Calendar action stages a date without
saving a revision. Selecting whole-report NOB stages seven NOB marks and an NOB
promotion recommendation. Schema 45 replaces unreferenced working snapshots after routine saves. Narrative
saves (main comments, employment, duties and Chief trait comments) and workflow
checkpoints are retained. Current heads keep unique revision identities for stale
write guards. Referenced snapshots remain protected for undo, review, signing,
conversation context and import previews; pre-policy/imported history is never
pruned. The editor history shows checkpoints and narrative saves. RevisionRetention is
linked by application checks because ReportRevision has reached the Access index
limit; migration rejects orphaned retention pointers. Comparisons reject replaced
working states and clear stale values rather than treating missing snapshots as
empty fields.
An active route also protects context when its writer changes, preserving field
ownership. Route flags use the fields actually changed in that save; higher,
grid and override contexts remain protected and hidden from checkpoint history.
An explicit clear of inherited native Block 28 records its actual prior text, even
when the older canonical snapshot was Null; later routine saves cannot revive
the native fallback. Imported provenance remains intact.

## 10. Testing and verification

- **Static checks before every build:** VBA source rules
  (`check-vba-source.ps1`), the build plan and test registry, canonical field
  definitions, and form and control references (every literal form and control
  name in the code must exist, with mutation checks proving a rename or removal
  fails).
- **Native tests** drive Access through COM against a fresh build. Each case
  runs on its own copy of the build, and any failure fails its suite. The
  build and test commands run on a separate hidden Windows desktop in the same
  sign-in session, so test windows cannot take input from the person using the
  workstation; nothing is installed and no setting is changed. Core
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
- **Dynamic view names are checked.** The source and reopened native build
  checks expand navigation, context-panel, detail-dialog, roster, senior and
  report-field control families from their catalogues. Native mutation tests
  exercise missing computed controls. New dynamic families need corresponding
  contracts; arbitrary computed strings are not proved safe by literal checks.
- **Receipts.** Each run writes a `verification.json` receipt under
  `build/workflow` with the build's hash. Receipts are local and not committed;
  verification results are recorded in commit messages.
- **Scratch retention.** Development and release workflows prune successful
  generated scratch artifacts after 24 hours and failed/flaky/unclassified
  scratch after seven days. Active sessions, pinned artifacts and supplied input
  databases are protected. Workflow/release receipts, development/release builds,
  historical fixtures and user workspaces are outside the cleanup scope. Local
  path/hash manifests record cleanup; deletion is not secure erasure. See
  [build scratch retention](docs/Development_Workflow.md#build-scratch-retention).

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
| 13 Nov 2026 | Development only, not in a release: the setup screen can write the blank copy's tables to a separate data file, or open an existing one, and then works through links (`modDataLink`, backlog S01). The data file holds tables only, no code. A missing or unusable data file is reported at startup with a file picker. A backup from a linked copy is a copy of the data file, verified row by row and field by field like any backup. Import and the session warning are not yet adapted to a linked copy. |
| 13 Nov 2026 | Internal, no change in behaviour: one module (`modDataFile`) answers where the workspace data is, and the backup source, file-size and location warnings, contributor session and table-level reads use it. First step toward a separate data file (backlog S01); the application is still a single file. The schema is defined in one file (`major-schema.json`). |
| 13 Nov 2026 | Editor Back Page: the comments and block 44 boxes have their full height in the fixed layout, so nothing overlaps them. Tests only: the import-engine suite seeds its source from the build under test instead of old release files, and the comment line capacities are asserted for the layout printer (Microsoft Print to PDF, 600 dpi). |
| 12 Nov 2026 | Development tooling only, no change to the application: build and test commands run their Access sessions on a hidden Windows desktop (section 10). |
| 10 Nov 2026 | Imports accepted from schema 35 (release 0.49.0) onward; the release gate tests release schemas only. Settings & Help > About this version and a 90-day Overview reminder, both offline. |
| 9 Nov 2026 | Release 0.55.0; full release gate and historical schemas 1-48 verified. |
| 9 Nov 2026 | Screen polish, no schema change: every form control names Segoe UI (Aptos defaults and theme fonts removed); buttons use two heights, 450 or 550 twips, except 14 where the taller size would overlap; a New button is the main (dark blue) action only when the screen has no available Open and no Save; the editor's context line takes two lines beside the report buttons; visible labels for the Reports due dismiss reason, member directory and snippet search, snippet text and My Work history; Overview recent reports show ISO dates; Summary groups buttons in sentence case. |
| 9 Nov 2026 | Schema 48: content checks can be turned off per match group (24, with stable IDs such as `RC-MEDICAL.conditions`). `RestrictedCheckOff` holds only the IDs that are off; no history is kept. The generator adds each group's label, the rule's plain-language reason and a regex-style display built from the same terms; the Content checks window has a Checks page with Turn on/off, Turn all off and Turn all on. |
| 9 Nov 2026 | Schema 47: restricted-content check. Reviewed rules in `src/restricted-content.json` (BUPERSINST 1610.10H 13-5 a-j and related paragraphs, plus application checks) generate `modRestrictedRules` via `scripts/generate-restricted-content.py`; `modRestrictedContent` matches whole words with prefixes, near and unless (directional `<`) terms, classification markings, identifiers and capitals in plain VBA, caches per block, and adds `RC-` findings to live validation. `RestrictedDismissal` (per report, rule, block and matched words) and `RestrictedException` (phrase, optional rule) hold dismissals and allowed phrases; both are journaled and reversible (Validation issues, Allowed phrases window). About 1 ms per 5,000-character pass. |
| 9 Nov 2026 | Line meter adds the PDF room left before 10 pt or an added page (`UXPrintHint`); group caption adds the EP and EP+MP spread from `PromotionCapacity` (`GroupPromotionSpread`); Reports due creates drafts and summary groups together (`CreateDueBatch`, all or nothing); the editor's Preview shows the report's PDF pages with pending edits (`UXEditorPrintPreview`). |
| 9 Nov 2026 | PDF comments (all three forms) and EVAL block 44 print at 12 pt, or at a reviewed 10 pt fallback (`fallbackSize` in `pdf-print-layout.json`, print row column 14) when the text does not fit at 12; only text that still does not fit goes to the added page. The editor's comments and block 44 boxes take the print box's font, width and height and follow the same 12/10 pt choice live. Block 44 prints below its caption (y 548.5, 30 pt). Measuring fonts are cached per size. Table sorting accepts value lists that end in a separator (Validation issues). |
| 9 Nov 2026 | Schema 46, field catalogue 3: EVAL block 44 (`Qualifications`) in the editor, NAVFIT98A import/export (`Qualifications` column), PDF import (XFA `BB`) and PDF output. Import of an earlier workspace restores block 44 on EVALs from the retained NAVFIT98A import value (`MigrateQualifications`) and records a report event. Retained NAVFIT98A values are now matched to the imported report by its native `ReportID` (`NativeReportMemo`); the legacy block 28 fallback previously read the row at the position equal to that ID. |
| 7 Nov 2026 | Release 0.54.0; full release gate and historical schemas 1-45 verified. |
| 7 Nov 2026 | Development 0.54.0: opening PDF output no longer uses `Shell.Application` (flagged by Navy macro threat detection); Access's `FollowHyperlink` opens it, the output path is shown to copy, and the source check rejects shell launches (section 6.7). |
| 6 Nov 2026 | Development 0.54.0: archiving ends open reviews; group reopen limited to completion sign-offs; operation journal forward-only transitions; revision-reference coverage test guards working-state deletion (sections 6.3 and 8). |
| 6 Nov 2026 | Development 0.54.0: one report lifecycle (`ReportLifecycle`) for group update/default guards; marking signed ends open routes and assignments with a recorded reason (section 6.3). |
| 6 Nov 2026 | Development 0.54.0: member profile saves update the member's open draft reports (skipping signed, approved, under-review, open and paygrade-conflicting ones); reporting senior fields and block 29 are no longer defaults; CHIEFEVAL-only PFA date stored as the first of its month; ISIC save is atomic. |
| 6 Nov 2026 | Development 0.54.0: startup location/session warnings and changed-contributor refusal (section 6.2), clock-safe recovery checkpoints (section 8), retained Members-grid review checkpoints and import undo from History (section 6.3). |
| 4 Nov 2026 | Schema 45: replaceable working report state, retained narrative/workflow checkpoints and a visible counselor picker. |
| 4 Nov 2026 | Development 0.54.0: senior-apply transaction ownership, NOB/date entry, source-match indication, guarded directory refresh and sibling report navigation (sections 9-10). |
| 4 Nov 2026 | Release 0.53.0; full release gate and historical schemas 1-44 verified. |
| 4 Nov 2026 | Schema 44: named report filters and import receipt attempts; counseling restore, ISIC import dependency order and dynamic view-family checks (sections 5.2, 9 and 10). |
| 4 Nov 2026 | Added automatic build/test scratch retention and local path/hash cleanup manifests (section 10). |
| 3 Nov 2026 | Schema 43: `DirectoryMergeLink` records roster history, departures and dismissals that a member merge moves, so Undo can move them back (section 5.2). |
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
