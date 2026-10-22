# Swift NAVFIT — current project state

Updated 22 October 2026. Current implementation, not a historical release log.
Consolidation removed three superseded release assemblers and consolidated
developer instructions, structural verification and Demo creation. Application
behavior is unchanged. The user-requested [PROPOSED_PURGE/](PROPOSED_PURGE/)
holds 91 historical builders/helpers/tests/browser files plus 2,567 generated VBA
exports pending review. Six reproducible bytecode caches were deleted; their
sources remain. The useful trait suite was restored and modernized. See the
[holding-area inventory](PROPOSED_PURGE/README.md)
for original paths, hashes and reasons. Unique old assertions remain available
for coverage/porting review; full replacement coverage is not claimed.
The [consolidation plan](planning/CONSOLIDATION_PLAN.md) records a completed bounded
review across all nine workstreams, including explicit retention decisions.
Review-held material and product backlog items are not approved for deletion or
represented as implemented redesigns.

## Purpose and supported behavior

Windows Microsoft Access application for drafting EVAL, CHIEFEVAL and FITREP
reports; command/member and prior-evaluation history; summary groups and RSCA;
revision comparison/restoration; pending-edit recovery; local routes/approvals;
shared report conversations; NAVFIT98A exchange and supported structured PDF import.
Local approvals are not official signatures. NAVFIT98A remains the native final
validation/printing destination. Official-template PDF output is a backlog item.

Current delivery is r128 / **0.27.1-dev / schema 27**. C03 contributor selection
is implemented: initialized launches require an explicit session contributor;
selection provides attribution, not authentication. C04 historical/effective-date
validation is implemented in development source; it is not a published release.
Git HEAD is `a4d8760`; r128 and environment changes remain local and uncommitted.

## Architecture and authoritative sources

- [src/](src/): VBA services, catalogues, form/macro templates. Version constants
  live in [modSwift.bas](src/modSwift.bas); major schema metadata lives in
  [major-schema.json](src/major-schema.json). Schema DDL also exists in the base
  builder and additive stages; there is not yet one complete schema authority.
- [build-manifest.json](scripts/build-manifest.json): 100 static modules, four
  generated modules, inputs and 72 ordered stages. [build.ps1](scripts/build.ps1)
  creates a fresh database, builds schema/forms, applies stages, compiles,
  reopens and verifies installed source and bindings.
- [AutoExec.txt](src/AutoExec.txt) invokes `SwiftAppStartup`: trust/setup/session
  contributor flow leads to the native application shell. Generated control
  expressions, form events and VBA dispatchers are runtime entry points too.
- Report editing uses an unbound buffer and fixed Front/Back geometry, seven
  traits on Front Page, a retained left rail, saved revisions and recovery.
  Window resizing changes the viewing area, not field geometry.
- [modMigration.bas](src/modMigration.bas) and
  [modWorkspaceImport.bas](src/modWorkspaceImport.bas) implement workspace import.
  Read source through read-only DAO without executing source code, verify exact
  original values/relationships before commit, and retain rollback/rejection guards.
- [fields.mjs](prototype/fields.mjs) is an **active native build input** consumed
  by [apply-ux-field-catalogue.ps1](scripts/apply-ux-field-catalogue.ps1).
  Browser-only code, its builder and four tests are held in PROPOSED_PURGE.
  This one metadata file remains at its original path to preserve the native
  input contract; browser storage and existing preview outputs are untouched.
- `frmMajorWorkbench` remains a summary-group context host; `frmBulkFolders`
  remains organizational-folder tooling. Their retention is explicit in the
  manifest. Old names alone do not establish redundant functionality.
- Unified report conversations retain legacy note/comment mappings and import
  history. Those tables and migration branches are not approved for removal.

## Environment, build, verification and delivery

Desktop Windows, Microsoft Access, Windows PowerShell 5.1 and Git are the current
development environment. No free Access Runtime, macOS/WSL application runtime,
ACCDE distribution or signing support is established. Python/Node are tooling
dependencies for particular helpers, not requirements of native PDF extraction.
Access Explorer and `swift_files` are already available; no new tooling is needed
for the consolidation. Use disposable copies for native investigation.

[develop.ps1](scripts/develop.ps1) is the standard fresh build/test entry point;
[tasks.json](.vscode/tasks.json) makes it the default build task. It also supports
Preflight, Build and Test actions. [release-check.ps1](scripts/release-check.ps1)
checks supplied Blank/Demo artifacts; it neither creates the Demo nor publishes.
[build-ux-admin-demo.ps1](scripts/build-ux-admin-demo.ps1) is the single supported
synthetic Demo builder: verified blank in, new seeded copy out, followed by
populated-Demo and startup checks. See the [Demo guide](planning/UX_DEMO_BUILD_GUIDE.md).
The older preview builder is quarantined; its distinct feature suites remain
available for change-specific verification instead of running on every seed.
See [Development_Workflow.md](docs/Development_Workflow.md) for current commands.

Artifact verification now reuses the builder's existing structural receipt only
for an unchanged artifact and complete matching input/checker fingerprint.
Metadata and requested blank inventories always run; workflow and migration
checks remain separate. Missing/stale receipts trigger native verification, and
`-ForceNative` requests it explicitly. No additional cache/dependency was added.

Deliver newly named Blank and synthetic Demo under [output/](output/), preserve
user workspaces, and verify uploaded names/sizes/folder membership in the designated
Drive folder. Close owned Access sessions normally, preserving pending user edits.

## Compatibility and verification limits

The importer currently accepts source schemas 1–27. Existing instructions promise
historical preservation; **the user's old-output testing waiver does not establish
that historical import support is abandoned**. r128 retested schema 27 transfer,
rollback and rejection guards, not schemas 1–26 or the populated schema-20 fixture.
The ongoing historical support/test policy remains unresolved. Do not restore
missing old outputs or narrow migration behavior by inference.

r128 native evidence is recorded in
[CONTRIBUTOR_STARTUP_R128_VERIFICATION.json](planning/CONTRIBUTOR_STARTUP_R128_VERIFICATION.json).
This inventory did not rerun it. Physical display, accessibility and printed fit
remain unverified. File/folder
permissions are the accepted protection boundary; masking does not encrypt SSNs.

### C04 validation applicability

The shared live/saved/group/export finding engine now selects CH-1 table 1 for
report ending dates from 16 December 2025 through 2026, and table 1A from
1 January 2027, including the O5/O6 transition and leap-year month ends.
Earlier calendars are unverified, not assumed invalid. General CH-2 policy
coverage begins 26 May 2026; missing dates and earlier reports receive an explicit
Needs review coverage finding, and unsupported policy assertions are downgraded.

Per the user's authority decision, NAVADMIN 132/26 controls the highest-paygrade
Regular periodic exemption from **3 June 2026**, not the CH-2 compilation date.
Permanent LDO grade and command-senior-enlisted eligibility still require human
confirmation; warnings never erase report fields. Findings retain policy references.

New saved runs select a new immutable applicability-v1 shared-readiness pack.
Existing packs, definitions, dates and results remain unchanged. The old r110
shared engine cannot be replayed under the updated implementation; select the
current pack for a new run. The separate bounded legacy draft pack is unchanged.
No schema, report geometry, migration or build-contract change is involved.
Verification on the fresh source build: [development workflow receipt](build/workflow/20261022-002149-1abcfadc-3df4-4471-9704-7e5d6362fd86/verification.json)
passes build/compile/reopen, installed source and bindings across 67 forms, blank
inventory and the standard workflow suites. The expanded
[readiness harness](tests/ReadinessHarness.bas) passes 150 assertions on a
disposable copy of that build; group correctness passes 39, rules 97 and retained
policy excerpts/page locators 47. Access Explorer confirmed native callback and
cross-object usage paths. Checks left the supplied fresh blank's hash unchanged.

The additional live-validation UI suite stops after 98 assertions at the existing
`UXValidationLayout` check. The same failure reproduces in an otherwise identical
disposable build with all four C04 modules restored to their prior Git versions;
it is not C04-specific. Later pagination/close assertions in that suite remain
unverified. No layout code or assertions were changed to suppress it. These are
development results, not release, physical-display or historical-import acceptance.

## Superseded approaches and open concerns

- The r126 local-PDF, r127 directory-merge and report-tools binary-template
  assemblers were removed in favor of the fresh manifest build. Their historical
  release evidence remains. Do not create another release assembler.
- Other historical patch assemblers and the duplicate preview builder are quarantined
  outside the active scripts directory. The build plan, canonical Demo builder and
  all manifest generators remain in place. Quarantined scripts must be restored
  to their original paths before any reviewed historical use.
- Browser synthetic behavior is not an alternative production implementation.
- The reviewed form layers remain ordered: base creation, metadata/surface,
  fixed pages, presentation, native unload guards and final route cleanup.
  Multiple passes are not sufficient evidence of redundant behavior. Legacy
  note mappings and separate audit histories retain different consumers.
- Mandatory conversation resolve/reopen and separate retired help/review screens
  must not be reintroduced; retained history and active adapters still matter.
- Full builds on every diagnostic edit and repeated verification of identical
  artifacts create avoidable delay. Consolidation must preserve real checks.
- Field/schema metadata and form generation are distributed; simplifying them
  is bounded follow-up work, not authorization for C05–C10 redesigns.
- The provenance viewer suite now creates synthetic receipts/raw values on its
  own disposable blank; it no longer requires the retired tools suite or an
  external database. Native import-capture fidelity remains separate. Unique
  assertions in the quarantined integrated tools suite still need a porting/
  coverage review; production provenance support is unchanged.
- Obsolete preview packaging and its separate schema-25 migration copy are held
  for review. The active workspace-upgrade tool imports canonical production
  migration source; no historical schema support was removed. Old diagnostic
  databases/logs remain protected even though the unused trace helper was moved.

The proposed batches and uncertain classifications are in the
[consolidation plan](planning/CONSOLIDATION_PLAN.md). Existing release notes are
historical evidence; this document owns the concise current architecture.
