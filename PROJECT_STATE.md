# Swift NAVFIT — current project state

Updated 22 October 2026. Current implementation, not a historical release log.
Consolidation removed three superseded release assemblers and consolidated
developer instructions, structural verification and Demo creation. Application
behavior was unchanged by consolidation. The persistence changes below are separate.
The user-requested [PROPOSED_PURGE/](PROPOSED_PURGE/)
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
Local approvals are not official signatures. NAVFIT98A remains supported for native
exchange and final validation. H01 development adds static official-template PDFs,
native page preview and combined summary-group packets; physical print acceptance
and a workstation test with Acrobat absent remain separate gates.

Current delivery is r128 / **0.27.1-dev / schema 27**. C03 contributor selection
is implemented: initialized launches require an explicit session contributor;
selection provides attribution, not authentication. C04 historical/effective-date
validation is implemented in development source. C05–C07 introduce
**0.28.0-dev / schema 28** persistence changes. C08–C10 extend development to
**0.29.0-dev / schema 29**. H01 advances application source to **0.30.0-dev**,
retaining schema 29; no new release is published.
M02 advances development to **0.31.0-dev / schema 30**, adding recoverable wizard
descriptors and explicit preview ownership. The earlier ACCDE candidate remains
unsigned 0.30.0-dev/schema 29 development evidence, not a current release.
M03 advances application source to **0.32.0-dev**, retaining schema 30. It batches
read work without changing storage, report geometry or the release contract.

## Architecture and authoritative sources

- [src/](src/): VBA services, catalogues, form/macro templates. Version constants
  live in [modSwift.bas](src/modSwift.bas); major schema metadata lives in
  [major-schema.json](src/major-schema.json). Schema DDL also exists in the base
  builder and additive stages; there is not yet one complete schema authority.
- [build-manifest.json](scripts/build-manifest.json): 105 static modules, five
  generated modules, inputs and 74 ordered stages. [build.ps1](scripts/build.ps1)
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
- [field-definitions.json](src/field-definitions.json) is the versioned report-field
  authority: storage, controls/recovery keys, labels, applicability, validation
  metadata, ordered NAVFIT/PDF bindings and fixed page geometry. The manifest
  fingerprints it. The native field stage generates the runtime catalogue;
  adapter/page modules delegate to it. [fields.mjs](prototype/fields.mjs) remains
  a retained browser reference, not the native catalogue authority. Compatibility
  JSON views and their generators derive from the canonical catalogue; the view
  generators no longer overwrite native VBA. Browser storage/outputs are untouched.
- `frmMajorWorkbench` remains a summary-group context host; `frmBulkFolders`
  remains organizational-folder tooling. Their retention is explicit in the
  manifest. Old names alone do not establish redundant functionality.
- Summary-group Members now has one presentation for newly created, empty and
  populated groups: the editable comparison grid in `frmUXComparison`, with
  member selection, shared-block changes, selected-report opening, validation,
  policy/roster review, reporting-senior/PSG RSCA, export and return actions below
  the table. `frmMajorWorkbench` retains Review/History and hidden legacy-tool
  context; its Members entry point forwards to the same live grid. Membership
  changes and report creation return to that grid, including removal of the last
  member. Empty grids derive trait captions from the group report type and
  disable opening a nonexistent selected report. Actions pass the displayed
  group explicitly and drain pending cells before opening another tool.
  The older group-read harness's deliberate `Report.Trait35` cache divergence
  fails the schema-29 persistence guard after 153 checks, identically on the
  tagged pre-change build and the unified-Members build. Its old expectation
  that this divergent cache remains readable was not weakened. Lifecycle,
  correctness and the focused Members/exit checks are verified separately.
  The [Members verification receipt](build/workflow/unified-members-85a9f12d-ef8d-4c07-aa44-8f4191310679/verification.json)
  records the [final fresh build](build/workflow/20261022-092532-83f31e78-a40f-4c0d-9713-70c0bb660fb9/verification.json),
  compile/reopen/source/binding and blank checks, and 118 final Members/exit
  assertions. The preceding complete feature build passed the standard suite,
  plus 42 lifecycle, 39 correctness and 57 contextual-help assertions. Its only
  subsequent source change was corrected empty-roster guidance, covered by the
  final focused suite. Empty and populated native previews confirm the combined
  table/action layout. One initial build lost its Access COM connection during
  theming; subsequent sequential fresh builds passed. No schema or dependency
  change, release or publication is included.
- Unified report conversations retain legacy note/comment mappings and import
  history. Those tables and migration branches are not approved for removal.
- [modPDFOutput.bas](src/modPDFOutput.bas) captures saved revisions or pending
  editor/grid values without applying them. [modPDFOutputUI.bas](src/modPDFOutputUI.bas)
  provides the export inventory, informational findings and native page preview.
  Pending group cells take precedence for the same field; unchanged grid-cache
  columns do not replace unrelated editor drafts, and both pending buffers survive.
  All listed group reports are sorted by last/first/middle name and stable identity;
  the summary and packet use that same order. Combined, individual and Both modes
  use one Access report, without a user-operated merger or Acrobat automation.
  Fixed Courier New values overlay embedded, reviewed 300-dpi blank templates.
  [pdf-print-layout.json](src/pdf-print-layout.json) records source/image hashes,
  print geometry and exclusions; it does not replace the canonical field catalogue.
  Unrecognized choice values and text beyond a printed box continue on labelled
  additional pages. A single-line value that cannot fit is marked `*` in its box,
  with its full value on an additional page. Private draft notes are excluded. Signatures and unrecorded
  qualifications/ranking/official RSCA/submission fields stay blank; computed PSG
  RSCA is not substituted for a report-scoped RSCA value. Individual exports use
  group naming/statistics only when membership identifies one group unambiguously.
  Folder preference is local to the Windows user and workspace identity. A chosen
  protected root receives a sanitized group folder and a unique dated run; staging
  remains `.partial` until every file and hash manifest succeeds. Failures retain
  their stage/error and protected partial output. Native preview stores only page
  ordinals in its disposable cache; captured personnel text stays in memory until
  PDF generation. Closing the flow removes its owned preview cache. A VBA reset
  invalidates the in-memory capture; refresh is required and an abandoned preview
  folder may require deliberate removal. Export does not change report revisions,
  status, approvals, recovery or signature dates. Operation outcomes use the existing
  journal. Runtime remains Access-only; PDF.js/pikepdf were development proof tools.
  Focused checks cover saved-field capture through the normal save service,
  coexisting pending editor/grid drafts, failed destinations, all three report types,
  native preview and a fully populated 144-member summary/290-page packet.
  Independent PDF text, page geometry, hash and image review are required alongside
  native assertions: text extraction alone cannot detect a white overlay masking
  data. CHIEFEVAL's narrow readiness-date and reporting-senior UIC slots use fixed
  7- and 9-point sizes; values are never automatically shrunk to fit. Native popup
  event tests pass, but popup visual acceptance remains unverified: Access Explorer
  captures the main Access window rather than that popup. Physical printed fit and
  operation on a workstation with Acrobat absent also remain unverified.
  The [final development receipt](build/workflow/20261022-134928-83b33732-1226-4436-a4b5-a9f611f325c3/verification.json)
  records a passing fresh BuildAndTest for 0.30.0-dev/schema 29. The unchanged
  artifact passed all 58 PDF assertions without replacement application imports;
  [independent structural checks](build/tests/pdf-output-e865466c-66c1-43a3-a84e-42b605e56b81/pdf-structure-qa.json)
  verified 15 PDFs/337 pages and all 11 file-hash manifests. These ignored local
  receipts are development evidence, not published release artifacts.
- [modAudit.bas](src/modAudit.bas) indexes retained domain evidence using stable
  event identities, immutable actor registrations, explicit correlation and
  affected-object foreign keys. Domain histories and undo remain authoritative.
- Six workspace routes in [modUX.bas](src/modUX.bas) feed rails and compact
  selectors. Retired tool aliases preserve editor context; required overview,
  report collection, group host and organizational folder tools remain.

## Environment, build, verification and delivery

### M02 preview ownership and recovery

`WizardPreview` is an additive schema-30 table. It stores the source/baseline
paths and SHA-256, owning contributor/workspace, target report/revision, reviewed
choice XML, a changing ownership token and save time. Values are rebuilt from
fingerprint-checked originals on deliberate **Resume saved preview**; they are
never applied automatically after a VBA reset or restart. XML parsing prohibits
DTDs/external resolution. Staged identifiers remain under the accepted C01 folder
permissions. No password, encryption or external dependency is introduced.

| Preview | Ownership and recovery policy |
| --- | --- |
| NAVFIT return | One named live context. Reviewed choices persist after each successful choice. Resume rechecks source, sealed baseline, contributor, mapping and target revision. Apply removes the exact owned stage in the normal import/revision transaction. |
| Structured PDF import | One named context owns extracted XML and the target snapshot. Reviewed century/identifier/merge mode and target persist. Resume re-extracts the original PDF and rebuilds the mapping. Apply remains explicit. |
| Whole-workspace import | An explicit descriptor owns path/hash and bootstrap workspace/profile stamp. Changed setup/source rejects Apply. Only the current import descriptor may occupy an otherwise blank target; its removal rolls back with failed import. Source preview rows are copied and verified with all other data. |
| Full/narrative NAVFIT import | Named bounded contexts bind source/hash, display name/century and contributor to the review form. Reset/restart requires a new preview. Existing source-reader cleanup remains in `CloseReviewFile`; a reset may leave an owned v33 conversion copy for deliberate cleanup. No temporary-file sweep is introduced. |
| Directory creation, membership/defaults, reviewed duplicate merge and bulk actions | Existing form selections and revision/member/group/default tokens reject missing or stale previews. Derived previews intentionally require re-preview after reset; domain history/undo/recovery remains authoritative. |
| PDF output, validation and parser caches | H01 retains its explicit captured job and invalidates it on reset. Captured pending report values are not staged by M02. Validation/parser caches are recomputed; the editor's durable recovery buffers remain unchanged. |

Closing a recoverable wizard retains its descriptor. Reopened screens require
Preview or Resume before enabling Apply; no source reader is retained by staging.
**Cancel saved preview** removes only the current contributor's preview of that
kind. Stale/damaged previews remain available for deliberate cancellation or
replacement. These ownership checks are attribution/staleness guards, not a new
authentication boundary. After a reset, close/reopen the wizard if former busy
controls remain disabled, confirm the contributor, then Resume.

The staging table references its target report/revision through foreign keys.
Ownership is checked by the service rather than adding a foreign key to
`UserProfile`, whose indexes already reach Access's limit; no existing relationship
or index was removed. Schema 1–29 import remains implemented. Focused development
proof covers actual VBE Reset, an intentional unhandled VBA failure, new-session
restart, changed source/target/contributor, cancel, late rollback/retry and retained
stage transfer. Schema-29/30 proof does not imply historical release acceptance.
The [fresh integrated development receipt](build/workflow/20261022-205637-f9035e36-ec54-4ab7-848c-cdcac94aa600/verification.json)
passes for 0.31.0-dev/schema 30. The [exact-artifact wizard receipt](build/tests/wizard-state-746e4b77-66ca-47ff-ac20-f1a155ad0382/verification.json)
records **63 passing assertions**, actual reset/unhandled-failure/restart proof,
filled-PDF Resume/Apply, and real schema-29 import with unchanged source hashes.
No replacement application source was imported during that final test. Static
native review covered all three Resume/Cancel rows; accessibility and physical
acceptance remain separate. The receipts are ignored local development evidence,
not release artifacts. No release, signing, publication or manifest change is included.

### M03 joined/batched read paths

| Path and inventory finding | Current contract |
| --- | --- |
| Reports / My Work | Existing `modUXReadRows` prepared presentation records remain in use. Route/group stamps are memoized within a refresh; write decisions still use their domain services. No competing display cache or persisted cache table was added. |
| Group/RSCA reads | `RSGroupRows` previously reopened head/cache/value/date/choice data per report. `AssertReportCacheRows` reads bounded 200-report chunks and shares projection, missing-field and typed-fact rules with the live scalar validator. A 500-report snapshot needs one report query plus 12 fact queries instead of roughly 2,501 validation/report reads, excluding constant RSCA context reads. No validation is skipped. |
| Bulk-folder tools | Selection stamps and Undo prechecks load selected report/move facts in 200-ID chunks. Folder labels use one refresh-owned folder map with lazy path resolution and the same missing-parent/depth errors. Input order, Null folder/revision markers and move IDs remain part of the stamp. `ApplyFolderBatch` and `AssignFolder` retain their existing live write checks, transactions, audit and guarded undo. |
| Directory history | Active prior traits and authoritative `ReportHead` revision traits are fetched in batches. Whole-report NOB, per-trait NOB, Null/blank/missing slots and incomplete prior records keep scalar semantics. The history list receives one completed value-list assignment instead of a native redraw per `AddItem`; the old query ordering and cell text remain unchanged. |
| Selected prior history / identity repair | Existing candidate/identity batching and per-refresh selected-source memoization remain. Identity matching/repair is a mutating domain boundary, not a presentation cache to bypass. Fresh save/permission/revision checks remain scalar where required. |

The folder-path map is replaced on `LoadBulkFolders`; VBA reset discards it and
the next display access rebuilds it. Nothing is reused for authorization or saved
state decisions. `FolderPath` and scalar cache validation remain available for
live writers. No framework, runtime dependency, schema change or manifest change
is part of M03.

The focused benchmark uses copies of the actual baseline/candidate artifacts,
the same 500-report/500-prior synthetic workspace, production import with original
field verification, complete output/order comparisons, five timed samples per
path, source hashes and no-read-history-write checks. Separate fault tests cover
541 assertions including typed facts, legacy missing values, nullable flags,
recovery, rollback and stale Undo. Timings are workstation measurements, not a
performance guarantee.
The [fresh integrated receipt](build/workflow/20261022-224259-661502dd-c337-4e14-bad2-e857db961fe3/verification.json)
and [exact-artifact benchmark/guard receipt](build/tests/batched-reads-c49c6563-bf32-4aef-8dca-e409e27482ba/verification.json)
pass for 0.32.0-dev/schema 30 without replacement source imports. At 500 reports
and 500 prior records, five-sample median milliseconds were:

| Read path | 0.31.0 baseline | 0.32.0 candidate |
| --- | ---: | ---: |
| Folder stamp | 903.75 | 16.77 |
| Folder display | 436.93 | 24.39 |
| Directory history | 22,357.69 | 7,616.27 |
| Group/RSCA | 2,210.68 | 677.01 |

The existing focused suites also pass: 34 persistence/typed-fact, 39 reporting
senior/RSCA, 72 directory/history, 21 directory UI, 31 cleanup/folder-history and
16 baseline/recovery assertions. Supplied artifacts and synthetic source hashes
remain unchanged. The receipts are local development evidence, not release or
publication records; no historical-output retesting was added by M03.

**Preexisting check gap:** on the untouched 0.31.0-dev baseline,
`UXReadRowsHarness` fails after assertion 77 when its deliberately diverged report
cache reaches guarded membership code (`Read saved member ... Trait33`). M03 does
not weaken that rejection or alter the old assertions. Its own scalar/batch tests
require identical cache/fact rejection codes. Historical-fixture, physical and
release acceptance remain separate.

### Workstation and compiled candidate support

Desktop Windows, Microsoft Access, Windows PowerShell 5.1 and Git are the current
development environment. No free Access Runtime, macOS/WSL application runtime,
production ACCDE distribution or signing acceptance is established. Python/Node are tooling
dependencies for particular helpers, not requirements of native PDF extraction.
Access Explorer and `swift_files` are already available; no new tooling is needed
for the consolidation. Use disposable copies for native investigation.

M01 adds an optional `Compile` action to [develop.ps1](scripts/develop.ps1),
converting a disposable copy of a verified current blank ACCDB into an unsigned
ACCDE development candidate. It leaves authoritative generation and the ACCDB
build contract intact. [test-compiled-distribution.ps1](scripts/test-compiled-distribution.ps1)
checks the actual compiled binary on disposable copies without injected VBA:
compiled status/source-export protection, all three report editors, save/reopen,
Settings import from ACCDB and populated ACCDE, retained raw import evidence,
history/provenance and unchanged source hashes. Receipts identify compiler version,
architecture, source/output hashes and checker fingerprints. Conversion uses the
undocumented Access SysCmd 603 action and requires exact-host verification; it is
not a portable supported API claim. The tested host is desktop Access
16.0.20430.20182, x64. No other architecture/version or Access Runtime is certified.

**M01 remains incomplete for production:** normal Office security on this host
displayed an unsigned-file security notice. After the probe cancelled that notice,
Access reported "Security settings restrict access to the file because it is not
digitally signed." This establishes an unresolved normal-policy launch gate, not
proof that every trusted location or explicit user trust decision would fail.
The owned dialogs were closed and the test owner quit; Trust Center policy was not changed. The current
user certificate store had no code-signing certificate. Trusted development
automation is not normal-policy launch acceptance. Compiled PDF generation also
remains unverified: external COM access to the compiled popup/control interface
failed, and that check was not counted as passed. H01's ACCDB checks remain valid.
The workspace picker still filters to ACCDB pending approval for the additional
application-source edit; direct production handlers successfully import ACCDE
data. Existing ACCDB Blank/Demo publication, historical import and release gates
remain required. [The workflow](docs/Development_Workflow.md#compiled-desktop-candidate-m01)
records compilation, signing/package integrity and remaining acceptance steps.
The [compiled development receipt](build/workflow/20261022-170507-27d28ca5-1cff-419a-ada3-49047c5dd297/verification.json)
records 34 passing trusted native assertions and the accepted candidate hash.
The overwrite guard also passed; the supplied ACCDB hash remained unchanged.

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

The importer currently accepts source schemas 1–28. Existing instructions promise
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

### C05–C07 persistence contracts

- `GroupCompletion` retains append-only sums/counts, senior and baseline lineage;
  `GroupCompletionReport` freezes each new counted report/revision and trait totals.
  `GroupCompletionCurrent` selects the current posting; `RSCountedReport` prevents
  duplicate accounting across groups. `GroupLifecycle` stores completion/reopen
  state separately from policy XML. Reopen/archive retains totals; correction
  appends a delta and supersedes active claims. Foreign keys guard deletion and
  revision ownership; counted groups cannot silently change reporting senior.
- Production import verifies every original field before additive backfill.
  Legacy XML and baseline notes remain unchanged as evidence; lineage becomes
  relational without adding a baseline or counting anything again. Legacy postings
  never recorded frozen revision IDs or per-report trait totals: these stay Null
  with explicit legacy evidence rather than being fabricated. Existing baseline
  history remains intact. New postings reconcile their exact child sums/counts.
- `ReportHead` owns the current revision with a composite ownership foreign key.
  The old circular Report-to-Revision foreign key is absent from the new target;
  the original current pointer and report columns remain checked query/interchange
  projections. Saved-field services read snapshots. Inconsistent projections stop
  a read/save rather than overwriting either value. Snapshot IDs, original field
  strings/nulls, change history and raw recovery buffers remain intact. New
  snapshots record their field-definition version; later catalogue fields can be
  absent from older snapshots without relabeling them.
- `RevisionDate` separates actual DATETIME values from Missing, Blank, counseling
  NOT REQ/NOT PERF and Unknown states for the five date/status fields. Original
  strings and counselor explanations remain in the immutable snapshot. Constrained
  `RevisionChoice` rows hold Retention/Statement responses, including explicit
  Conflict/Unknown review states; native Yes/No boxes remain lossless projections.
  Deliberate selection clears the opposite box. The single-choice commit API
  accepts Yes/No/Unspecified. Form/workflow domains reject unsupported new values
  while retaining exact imported legacy exceptions. `FieldDefaultLink` verifies
  one scope target through real foreign keys; unresolved legacy links are retained
  and cannot be applied as defaults.

Focused checks cover relational posting/migration/rollback, typed/null/raw values,
snapshot authority and a metadata-only synthetic field through save, restore and
fresh-session recovery. Populated real-schema-27 and current-schema synthetic
sources pass Settings import, exact source-field verification, late rollback,
retry, idempotency and reopen. Schemas 1–26 were not retested for this change;
the historical-fixture waiver/support distinction above still applies.
The [final development receipt](build/workflow/20261022-015423-b7d4be0c-ca65-4f69-b030-55257beeca5f/verification.json)
passes fresh build/compile/reopen, installed source/bindings across 102 modules
and 67 forms, blank inventory and standard workflow suites. The final blank's
SHA-256 is `72CF2A7E011D36797E114799B7A20BA1F546D26505405BC731BE745D4D0173FC`.
The [focused logs](build/workflow/c05-c07-20261022-020443-3da1242d-7b0d-4f80-b3fe-9c77c7108690/)
record 34 persistence, 43 accounting, 569 catalogue, 200 default, 507 NAVFIT
adapter, 410 exchange, 99 interchange-guard, 12 PDF counseling and 16 engine
assertions, plus populated schema-27/28 transfer and late rollback.
The [boundary receipt](build/workflow/c05-c07-final-20261022-022006-63fe221a-e2cb-4b97-8b99-54815e9d6660/verification.json)
links the additional catalogue text/date extension and older-version snapshot
preservation check, 170 service assertions/49 reference checks, and passing
decision (198), report-people (364), comparison (22), traits (54) and structured-name
(281) suites with fresh-session recovery. The supplied blank remained unchanged.
The compatibility-view generators also reproduce all four JSON reference views
in temporary storage using only the existing Python standard library.

Earlier field-count assumptions in these feature harnesses now use the canonical
field list; new test sessions explicitly select their synthetic contributor.
Their original layout assertions remain available. Decisions, report people and
structured names were run with explicit `-PersistenceOnly`: their old button-bound
checks reference buttons also absent on the unchanged schema-27 baseline.
The full RSCompletion suite's separate `Draft notes hidden` assertion likewise
reproduces on that baseline; `-AccountingOnly` retains all accounting assertions.
Current editor geometry/pending behavior passes 264 assertions in the standard
fixed-page suite. These legacy UI gaps and the separate C04 layout gap remain;
no release, physical display, accessibility or printed acceptance is implied.

## C08–C10 audit, conversations and navigation

Development schema 29 adds AuditActor/AuditEvent without removing earlier tables.
Stable event/source identities, actor/time, explicit operation correlation and
affected-object links index distinct revision, report-event, folder, review,
conversation and operation histories. Operations displays this timeline. Start
and outcome envelopes preserve the original operation author and the actual
completing actor separately. Unknown historical context stays unknown; existing
undo and original histories remain. The
[producer/consumer map](planning/CONSOLIDATION_PLAN.md#c08c10-producer-consumer-and-navigation-audit)
records the retained consumers and precise generation/dispatch ownership.

Current posted notes/replies use ReportMessage. Legacy posting adapters delegate
to that service; new results use negative canonical message tokens to avoid
colliding with positive legacy IDs. Private mutable checkpoints retain the
ReportNoteDraft storage and author/version guards. Legacy rows, mappings,
timestamps, authors and ancestry remain importable; conversion allocates parent
identities before linking, verifies existing mappings and rejects invalid
ancestry atomically. Normal conversation/inbox reads no longer perform conversion.
Posting remains independent of approval, with no new resolve/reopen workflow.
Schema-29 import does not implicitly publish DraftNotes into shared messages.

One six-route list feeds rails and compact selectors. Fresh generation no longer
constructs the retired help/checklist/comments/composer or temporary review
screen; it creates frmReportReviewPanel directly. Compatibility tool aliases,
contextual help, required group host and folder tooling remain. Native form and
macro exports capture actual callbacks and source bindings; fixed pages and
pending-edit guards remain unchanged. A retained schema-21/22 directory-backfill
read now uses the owning DAO transaction, avoiding a Workspace UI-connection lock.

No dependencies were added. The two new focused runners are
[audit/conversation/navigation](scripts/test-audit-conversations-navigation.ps1)
and [retained-schema transfer](scripts/test-audit-message-transfer.ps1).
The NAVFIT UI runner now points to the retained reference template rather than
its obsolete root-folder path; the original harness assertions are unchanged.
The [final fresh build receipt](build/workflow/20261022-081737-c354c3e3-1221-43ea-845a-36ed661924d3/verification.json)
passes all 72 stages, compile/reopen, installed source/bindings (103 static
modules, 67 forms), blank inventory and the standard development suite.
SHA-256: `F8B6883CCFE9DEABF9DE2C811DD19DFEACDD2C6F54B91C5B92B946C5E93994B9`.
The [focused receipt](build/workflow/c08-c10-final-20261022-0827/verification.json)
links 1,988 audit/conversation/navigation, contextual-help, My Work, group and
cleanup assertions on that exact artifact, plus the native form/macro inventory.
The [transfer matrix](build/tests/audit-message-transfer-fa66cd5c-a0e3-4544-8acc-49d7218a7502/verification.json)
passes actual retained schemas 21–29, including a populated schema-22 source
and populated schema-29 audit history. Original fields/IDs/history are verified
before commit; later-numbered parents, private reply context, idempotency,
close/reopen and original/source/template hashes pass. A forced late audit
failure restores bootstrap data and rolls back all imported domain/audit rows;
retry succeeds. Pre-final schema-29 regression checks additionally pass 34
persistence, 43 accounting, 507 adapter, 410 exchange, 99 exchange-guard, 35
NAVFIT UI, 12 PDF counseling and 16 engine assertions plus adversarial import
rollback/retry and future-schema/unknown-table/changed-source/populated-target
rejection. The final production-source delta after those regression checks is
the transaction-owned schema-21/22 directory read, exercised by the final matrix.
Schemas 1–20 are still supported by migration branches but were not retested;
the recorded old-output waiver remains, and no missing fixtures were recreated.
Existing C04 and legacy feature-layout gaps above remain separate from these
checks. No release, publication, physical/accessibility or printed acceptance is
claimed.

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
- Non-report schema metadata and form-generation passes remain distributed.
  C08–C10 retain these boundaries; unrelated consolidation remains out of scope.
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
