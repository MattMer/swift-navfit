# Swift NAVFIT technical README

This document is for technical reviewers, command ISSMs, and developers. It
covers how Swift NAVFIT is built, how it handles data, how it is tested and
released, and what is not yet verified. For what the application does and how
to use it, see the [README](README.md) and the
[user guide](docs/Swift_NAVFIT_User_Guide.md).

## At a glance

| | |
| --- | --- |
| Platform | Windows, full desktop Microsoft Access (64-bit tested). The free Access Runtime is not supported. |
| Tested host | Access 16.0.20430.20182 x64. |
| Language | VBA, with Access forms, queries and one macro (`AutoExec`). |
| Deployment | One `.accdb` file per command workspace. No installer, server or service. |
| Network | None. No telemetry, update check, sync or web service. |
| Third-party runtime code | None. Everything called at runtime is part of Windows or Microsoft Office (see the [SBOM](docs/architecture/SBOM.md)). |
| Latest release | 0.50.0, schema 35 (31 October 2026), tag `v0.50.0`. |
| Development source | 0.51.0-dev, schema 36. `APP_VERSION` and `SCHEMA_VERSION` are in [modSwift.bas](src/modSwift.bas). |
| Source size | 106 VBA modules, 70 forms as text, JSON catalogues, 71 tables. |
| Tests | 33 test suites over 120 case scripts and 115 VBA test harnesses, plus static source checks before every build. |
| Signing | Not signed yet. Access shows the Enable Content bar. |

## Architecture

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
  contracts. A build takes about 2 minutes. Each build writes a
  `.build.json` receipt that fingerprints every input.
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
  lifecycle states are stored as rows with foreign keys, so RSCA totals can be
  reconciled by query.
- **Audit and history:** `AuditEvent` indexes the domain histories (revisions,
  review decisions, merges, sign-offs, draft disposals). Old audit, note and
  comment records are kept and stay importable.

The [architecture description](docs/architecture/Swift_NAVFIT_Architecture_Description.md)
and [PROJECT_STATE.md](PROJECT_STATE.md) give the full picture, including the
report editor, navigation and PDF output.

### Repository layout

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

## Data handling and security

- **Storage.** All report data, history, review records and recovery
  checkpoints are in the workspace `.accdb`. Protection relies on the Windows
  permissions of the folder it lives in (a deliberate decision, recorded as C01
  in the [backlog](planning/PRIORITY_BACKLOG.md)): there is no database password
  or encryption.
- **SSNs** are stored readable in current values, revisions, recovery
  checkpoints and import provenance. Masking in lists and previews is display
  only, and SSNs are kept out of generated file names.
- **Identity.** Contributor profiles record who made each change. They are
  attribution, not authentication; anyone who can open the file can choose any
  contributor.
- **No network use.** Runtime external calls are limited to MSXML 6,
  `Scripting.FileSystemObject`, `ADODB.Stream`, and Windows GDI and crypto API
  declarations. Swift starts no child processes. Optional reference links open
  the default browser only when selected and carry no report data.
- **Untrusted input.** Imports (NAVFIT98A databases, earlier Swift workspaces,
  PDFs, RUAD CSV) are read as data:
  - earlier workspaces are opened through read-only DAO and their code is never
    run;
  - every XML parser disables DTDs and external resolution, and a source check
    fails any parser that doesn't;
  - `test-malformed-imports` runs 55 malformed and control inputs against the
    PDF, CSV and NAVFIT98A importers;
  - imports preview before applying, and a failed import rolls back.
- **Startup lockdown is presentation, not protection.** Builds open as a plain
  application window with the navigation pane, full menus and ribbon hidden.
  Holding Shift while opening bypasses this (`AllowBypassKey` stays on by
  decision, for support and investigation).
- **Macro trust.** Because the files are unsigned, users must choose Enable
  Content, and a Navy-managed workstation's Office policy may block that. The
  [code-signing plan](docs/architecture/Threat_Model.md#5-code-signing-plan-m01)
  needs a DoD PKI code-signing certificate.
- **Open threat-model recommendations:** sign releases; optionally hash each
  saved revision to detect direct table edits. See the
  [threat model](docs/architecture/Threat_Model.md).

## Data integrity and recovery

- Saves run in a transaction; a failed save rolls back and keeps the pending
  values for retry.
- Unsaved edits are checkpointed separately from saved revisions, so they
  survive page changes, failed saves, a VBA reset and a restart.
- Restoring an earlier version writes a new revision; history is never
  rewritten.
- Signed reports refuse every save path until reopened, and draft disposal keeps
  the revision rows so every link stays valid.
- Backups are data-only `.accdb` copies written and verified inside Access.
  Restore is an import into a new blank workspace.

## Compatibility and upgrades

- The importer accepts source workspaces of schema 1 through 36, into a target
  of schema 21 through 36. It refuses a populated target, an unknown source
  table, or a source newer than the application.
- Every value is compared with the source before the import commits.
- Every release runs the workspace-migration suite over every schema from 1 to
  the current one. [tests/historical-schemas](tests/historical-schemas/README.md)
  holds each older schema's DDL: 8–20 exact from Git, 21–32 captured from real
  builds, 1–7 reconstructed (they predate Git; undocumented column changes there
  are untested).

## Testing and verification

- **Static checks before every build:** VBA source rules
  (`check-vba-source.ps1`), the build plan and test registry, canonical field
  definitions, and form and control references (every literal form and control
  name in the code must exist, with mutation checks proving a rename or removal
  fails).
- **Native tests** drive Access through COM against a fresh build. Each case
  runs on its own copy of the build, and any failure fails its suite.

  | Tier | Suites | When |
  | --- | --- | --- |
  | Core | 8 | Every build (`develop.ps1 -Action BuildAndTest`, about 5 minutes). |
  | Extended | 19 | Before a release or a risky change (`-Action ExtendedTest`, about an hour). |
  | Release | 1 | Release candidates only (`release-check.ps1`). |
  | Helper and manual | 5 | Fixtures, benchmarks and attended checks. |

- **Coverage includes** persistence and rollback, recovery after a simulated
  crash, imports of every historical schema, NAVFIT98A round trips, PDF output
  fixtures, validation rule packs, responsive layout (over 900 assertions),
  accessibility names and contrast, and the architecture views.
- **Generated documents are checked.** Architecture views, the SBOM and state
  lists are regenerated from a build, and a test fails if they drift from it.
- **Receipts.** Each run writes a `verification.json` receipt under
  `build/workflow` with the build's hash. Receipts are local and not committed;
  verification results are recorded in commit messages.

## Release process

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
published hashes identify the released files.

## Performance and limits

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

## Known limitations and open items

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
- **Intermittent test failures, root cause unknown:** one Jet error 3802 in a
  review-assignment audit write, and one DAO error 3059 in the import timing
  test. Both passed on retry and in repeated standalone runs; a retry now keeps
  the first failure's log.
- **Accessibility:** automated checks pass with no known gaps, but there is no
  screen-reader acceptance and no formal Section 508 conformance report yet.

The full list is in [PROJECT_STATE.md](PROJECT_STATE.md#known-gaps-and-test-caveats)
and the [priority backlog](planning/PRIORITY_BACKLOG.md).

## Architecture and compliance documents

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

## Building from source

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

## Further reading

- [Current project state](PROJECT_STATE.md) and its [history](planning/PROJECT_STATE_HISTORY.md)
- [Priority backlog](planning/PRIORITY_BACKLOG.md), [recent development](planning/RECENT_DEVELOPMENT.md) and [release history](planning/CURRENT_STATE.md)
- [Build and verification history](planning/BUILD_STATUS.md)
- [Summary groups, review routes and workflow foundations](planning/MAJOR_FOUNDATIONS_GUIDE.md)
- [Workspace guide](planning/UX_WORKSPACE_GUIDE.md)
- [Review assignments and My Work](planning/REVIEW_ROUTING_GUIDE.md)
- [Report tools, unified notes and verification](planning/REPORT_TOOLS_R123_2026-10-20.md)
- [Reporting seniors and averages](planning/REPORTING_SENIOR_GUIDE.md)
- [Workspace transfer and fixed report pages](planning/WORKSPACE_COMPAT_PAGES_R115_2026-10-17.md)
- [Import provenance](planning/PROVENANCE_GUIDE.md)
- [Historical review-comment service notes](planning/REVIEW_COMMENTS_GUIDE.md)

## License

This repository has no open-source license. Contact its owner about use and
redistribution.
