> **Historical record.** This is the former PROJECT_STATE.md as of 24 October 2026 (0.44.0-dev / schema 33), kept for its per-feature implementation detail and verification evidence. Links into `build/` point to ignored local receipts. The current state is in [PROJECT_STATE.md](../PROJECT_STATE.md).

# Swift NAVFIT — project state history (through 0.44.0-dev)

Updated 24 October 2026. Current implementation, not a historical release log.
Consolidation removed three superseded release assemblers and consolidated
developer instructions, structural verification and Demo creation. Application
behavior was unchanged by consolidation. The persistence changes below are separate.
The user-requested [PROPOSED_PURGE/](../PROPOSED_PURGE/)
holds 91 historical builders/helpers/tests/browser files plus 2,567 generated VBA
exports pending review. Six reproducible bytecode caches were deleted; their
sources remain. The useful trait suite was restored and modernized. See the
[holding-area inventory](../PROPOSED_PURGE/README.md)
for original paths, hashes and reasons. Unique old assertions remain available
for coverage/porting review; full replacement coverage is not claimed.
The [consolidation plan](CONSOLIDATION_PLAN.md) records a completed bounded
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
M04–M06 advance development source to **0.33.0-dev / schema 31**: contextual
diagnostics, required view contracts and versioned runtime trait reference data.
H02 advances development source to **0.34.0-dev / schema 31**, allowing saved
report/group NAVFIT98A export despite readiness findings and archived/approved
status. H01's corrected Chief fields and rendering advance source below;
physical and renderer-specific acceptance remain separate.
L01–L02 advance development to **0.35.0-dev / schema 32**: descriptive major
relationship names and writable revision-ID compatibility queries. Original
stored column names and revision IDs remain intact for historical import.
H01 corrections advance development to **0.36.0-dev / schema 33**, with field
catalogue version 2 and 21 independent authored Chief trait comments.
Remote regression and accessibility preparation advances development to
**0.37.0-dev / schema 33**. Saved-value export entry points preserve pending
buffers, comparison projections tolerate other canonical proposals, and entry/
validation controls have explicit label associations. No storage change is added.
Reliability follow-up advances development to **0.38.0-dev / schema 33**:
current-schema ACCDB/ACCDE backup/restore and stale-preview cancellation protection.
M02/M03/M04 follow-ups advance source to **0.39.0-dev / schema 33**: workspace
preview reopen/cancellation ownership, guarded My Work refresh ownership and
primary export diagnostics with truthful file-publication outcomes.
M04/M05 follow-ups advance source to **0.40.0-dev / schema 33**, distinguishing
committed workspace imports from confirmed rollback and strengthening required
wizard controls/action bindings. No storage or report-layout change is added.
Further M04/M05 work advances source to **0.41.0-dev / schema 33**, with typed
PDF/return Apply outcomes and stronger Members/My Work view contracts.
User-reported UI corrections advance source to **0.42.0-dev / schema 33**.
Embedded conversation sorting resolves subforms through their actual hosts.
Report tools replaces the redundant Notes button. Chief trait comments use the
official comment-box dimensions and typography, with read-only selected-grade
wording beneath each trait title. All seven traits remain on the editing Front
Page; printed page associations are unchanged. Editor recovery identity, original
revision and return-group context survive VBA reset. When an automatic save
fails, navigation may proceed only after a successful raw-edit checkpoint;
stale saves remain rejected and recovery failure keeps the editor open.
The summary list offers selected-group NAVFIT98A export and PDF preview/output.
The NAVFIT export screen requires only the selected group and output directory;
the reviewed native blank is embedded as a generated resource and hash-verified
before use. Legacy return services, baseline/provenance and saved recovery data
remain intact, without visible return/template/report pickers. The duplicate
Members header Back button and Reports Folder tools screen are retired; folder
records, history, migration and backend services remain supported.
M02/M04/M05 follow-ups advance source to **0.43.0-dev / schema 33**: native-group
import, move-to-group and duplicate-member merge previews belong to the reviewing
contributor. Contributor changes, failed Apply, reopen and VBA reset require a
fresh preview. Native group import exposes its transaction outcome to the UI;
committed/display-failure and uncertain outcomes cannot claim rollback or retain
an actionable preview. Completed moves/merges remain explicit after display
failure. All three screens have required control/action contracts. No schema,
report geometry, external dependency or release change is included.
M02/M03 follow-ups advance source to **0.44.0-dev / schema 33**. Directory-batch
report creation binds the reviewing contributor, invalidates stale state before
every preview and on failed Apply/close/reopen/filter changes, and exposes explicit
transaction outcomes to its UI. Invalid dates cannot leave an older batch
actionable; committed/display-failure feedback cannot claim nothing was saved.
The summary-group list joins the unique lifecycle row for displayed workflow
status instead of repeated scalar lookups. Counts, closeout/default semantics,
filters, ordering and live write guards remain unchanged.

## Architecture and authoritative sources

- [src/](../src/): VBA services, catalogues, form/macro templates. Version constants
  live in [modSwift.bas](../src/modSwift.bas); major schema metadata lives in
  [major-schema.json](../src/major-schema.json). Schema DDL also exists in the base
  builder and additive stages; there is not yet one complete schema authority.
- [build-manifest.json](../scripts/build-manifest.json): 105 static modules, six
  generated modules, inputs and 74 ordered stages. [build.ps1](../scripts/build.ps1)
  creates a fresh database, builds schema/forms, applies stages, compiles,
  reopens and verifies installed source and bindings.
- [AutoExec.txt](../src/AutoExec.txt) invokes `SwiftAppStartup`: trust/setup/session
  contributor flow leads to the native application shell. Generated control
  expressions, form events and VBA dispatchers are runtime entry points too.
- Report editing uses an unbound buffer and fixed Front/Back geometry, seven
  traits on Front Page, a retained left rail, saved revisions and recovery.
  Window resizing changes the viewing area, not field geometry.
- [modMigration.bas](../src/modMigration.bas) and
  [modWorkspaceImport.bas](../src/modWorkspaceImport.bas) implement workspace import.
  Read source through read-only DAO without executing source code, verify exact
  original values/relationships before commit, and retain rollback/rejection guards.
- [field-definitions.json](../src/field-definitions.json) is the versioned report-field
  authority: storage, controls/recovery keys, labels, applicability, validation
  metadata, ordered NAVFIT/PDF bindings and fixed page geometry. The manifest
  fingerprints it. The native field stage generates the runtime catalogue;
  adapter/page modules delegate to it. [fields.mjs](../prototype/fields.mjs) remains
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
  The [Members verification receipt](../build/workflow/unified-members-85a9f12d-ef8d-4c07-aa44-8f4191310679/verification.json)
  records the [final fresh build](../build/workflow/20261022-092532-83f31e78-a40f-4c0d-9713-70c0bb660fb9/verification.json),
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
- [modPDFOutput.bas](../src/modPDFOutput.bas) captures saved revisions or pending
  editor/grid values without applying them. [modPDFOutputUI.bas](../src/modPDFOutputUI.bas)
  provides the export inventory, informational findings and native page preview.
  Pending group cells take precedence for the same field; unchanged grid-cache
  columns do not replace unrelated editor drafts, and both pending buffers survive.
  All listed group reports are sorted by last/first/middle name and stable identity;
  the summary and packet use that same order. Combined, individual and Both modes
  use one Access report, without a user-operated merger or Acrobat automation.
  Field-specific fonts overlay embedded, reviewed 300-dpi blank templates;
  Windows printer GDI metrics drive width, wrapping and line height.
  [pdf-print-layout.json](../src/pdf-print-layout.json) records source/image hashes,
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
  The [final development receipt](../build/workflow/20261022-134928-83b33732-1226-4436-a4b5-a9f611f325c3/verification.json)
  records a passing fresh BuildAndTest for 0.30.0-dev/schema 29. The unchanged
  artifact passed all 58 PDF assertions without replacement application imports;
  [independent structural checks](../build/tests/pdf-output-e865466c-66c1-43a3-a84e-42b605e56b81/pdf-structure-qa.json)
  verified 15 PDFs/337 pages and all 11 file-hash manifests. These ignored local
  receipts are development evidence, not published release artifacts.
- [modAudit.bas](../src/modAudit.bas) indexes retained domain evidence using stable
  event identities, immutable actor registrations, explicit correlation and
  affected-object foreign keys. Domain histories and undo remain authoritative.
- Six workspace routes in [modUX.bas](../src/modUX.bas) feed rails and compact
  selectors. Retired tool aliases preserve editor context; required overview,
  report collection and group host remain. The obsolete Folder tools screen is
  retired; its legacy navigation alias returns to Reports without changing data.

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
| Selected NAVFIT group import | Source hash, roster selection and reviewing contributor bind the in-session preview. Failed Apply clears live ownership; a fresh preview is required. The UI distinguishes confirmed rollback, committed/display failure and uncertain outcomes. No new recovery table or automatic replay is introduced. |
| Directory creation, membership/defaults, reviewed duplicate merge and bulk actions | Existing form selections and revision/member/group/default tokens reject missing or stale previews. Derived previews intentionally require re-preview after reset; domain history/undo/recovery remains authoritative. |
| PDF output, validation and parser caches | H01 retains its explicit captured job and invalidates it on reset. Captured pending report values are not staged by M02. Validation/parser caches are recomputed; the editor's durable recovery buffers remain unchanged. |

Move-to-group and duplicate-merge UI previews additionally bind the reviewing
contributor in 0.43. Reopening the screen clears the token; failed or cross-owner
Apply clears the token and disables Apply. Live domain revision, identity,
membership, archive, approval, atomicity and guarded-undo checks remain unchanged.
In 0.44, directory-batch creation applies the same contributor boundary in its
domain service. Failed preview clears the old context before validating new
inputs. UI opening/closing/filter changes and failed Apply invalidate the context.
Failed validation leaves visible member/date inputs available for correction;
reopening the screen requires choosing and previewing its inputs again.
The context is derived/in-session; no new staging table or automatic replay exists.

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
The [fresh integrated development receipt](../build/workflow/20261022-205637-f9035e36-ec54-4ab7-848c-cdcac94aa600/verification.json)
passes for 0.31.0-dev/schema 30. The [exact-artifact wizard receipt](../build/tests/wizard-state-746e4b77-66ca-47ff-ac20-f1a155ad0382/verification.json)
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
| Summary-group list | In 0.44, the existing aggregate query directly joins the primary-key lifecycle row for workflow display. Missing/blank workflow still displays Initialized; member counts, archived filters, empty-group default dates and mixed/incomplete closeout text remain unchanged. `GroupWorkflow` remains available for live domain guards and other callers. |
| Group/RSCA reads | `RSGroupRows` previously reopened head/cache/value/date/choice data per report. `AssertReportCacheRows` reads bounded 200-report chunks and shares projection, missing-field and typed-fact rules with the live scalar validator. A 500-report snapshot needs one report query plus 12 fact queries instead of roughly 2,501 validation/report reads, excluding constant RSCA context reads. No validation is skipped. |
| Retained folder services | Selection stamps and Undo prechecks load selected report/move facts in 200-ID chunks. Input order, Null folder/revision markers and move IDs remain part of the stamp. `ApplyFolderBatch` and `AssignFolder` retain live write checks, transactions, audit and guarded undo. The former folder-screen display benchmark is historical; its screen and mutation UI callbacks are retired in 0.42. |
| Directory history | Active prior traits and authoritative `ReportHead` revision traits are fetched in batches. Whole-report NOB, per-trait NOB, Null/blank/missing slots and incomplete prior records keep scalar semantics. The history list receives one completed value-list assignment instead of a native redraw per `AddItem`; the old query ordering and cell text remain unchanged. |
| Selected prior history / identity repair | Existing candidate/identity batching and per-refresh selected-source memoization remain. Identity matching/repair is a mutating domain boundary, not a presentation cache to bypass. Fresh save/permission/revision checks remain scalar where required. |

The former folder display map/benchmark is historical after the 0.42 screen
retirement; retained folder batch/history services remain supported. Nothing is
reused for authorization or saved-state decisions. `FolderPath` and scalar cache
validation remain available for live writers. No framework, runtime dependency,
schema change or manifest change is part of M03.

The focused benchmark uses copies of the actual baseline/candidate artifacts,
the same 500-report/500-prior synthetic workspace, production import with original
field verification, complete output/order comparisons, five timed samples per
path, source hashes and no-read-history-write checks. Separate fault tests cover
541 assertions including typed facts, legacy missing values, nullable flags,
recovery, rollback and stale Undo. Timings are workstation measurements, not a
performance guarantee.
The [fresh integrated receipt](../build/workflow/20261022-224259-661502dd-c337-4e14-bad2-e857db961fe3/verification.json)
and [exact-artifact benchmark/guard receipt](../build/tests/batched-reads-c49c6563-bf32-4aef-8dca-e409e27482ba/verification.json)
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

**Earlier check gap, repaired in 0.38:** on the untouched 0.31.0-dev baseline,
`UXReadRowsHarness` fails after assertion 77 when its deliberately diverged report
cache reaches guarded membership code (`Read saved member ... Trait33`). M03 does
not weaken that rejection or alter the old assertions. The repaired fixture
explicitly verifies rejection and no membership write, restores its exact synthetic
values, then continues the remaining review/filter/sort assertions. Scalar/batch tests
still require identical cache/fact rejection codes. Historical-fixture, physical and
release acceptance remain separate.

### Workstation and compiled candidate support

### H01 Chief comments and print-fidelity corrections

- Field catalogue version 2 appends 21 Chief-only memo values, named
  `Trait33Comment1` through `Trait39Comment3`, after the original 79 ordinals.
  Each trait has three independent entries beside its existing Front Page row.
  All 188 existing field rectangles remain fixed; the Chief guidance panel moves
  below the seven groups. Other editor kinds and printed page associations stay
  unchanged. Rating selection never authors or replaces performance comments.
- Schema 33 creates the additional report/recovery columns through the existing
  canonical builder. Normal snapshots, revisions/changes, comparison, restore,
  pending buffers, shutdown recovery and workspace import consume the catalogue.
  The base form builder creates appended unbound buffers before shared layout;
  the existing surface generator creates their visible controls.
  Old snapshots remain unchanged and their new fields stay null/not recorded;
  missing fields in a snapshot that claims the new catalogue remain corruption.
- Chief native mappings use `LEAD/TAC/PROF/MIS/EO/TEAM/MIL` plus `DN1`–`DN3`.
  Read-only DAO verified each native column is text(255); retained loader IL matches
  the installed NAVFIT98A binary SHA256 and establishes the control/column roles.
  There are now 96 Chief mappings and 75 EVAL/FITREP mappings; all 128 native
  columns/provenance remain covered. Structured-PDF input maps the 21 corresponding
  `FONTB1`–`FONTB3` fields independently. Native serialization capacity still fails
  truthfully; full memo text and PDF continuation are never truncated to 255.
- Prior Swift imports held these columns only in native provenance. An untouched
  null new value does not erase that retained native text. Once authored or
  deliberately cleared, mapped writes apply normally. Return reconciliation uses
  the actual sealed native comment baseline, including older SwiftValue tables
  without these new keys. Unchanged native text does not manufacture authored
  history; a changed legacy value requires the appropriate three-way decision.
- Printed Chief grade headings, full catalogue explanations and boxed numeric
  marks occupy their own areas. Authored text alone fills the 21 performance
  comment areas. Comments use the template's Times New Roman 12-point style;
  grade wording normally uses 7 points. The official trait-34 grade-4/5 style is
  6.5 points, which native Access plain and rich text both round. Development
  uses a documented 6-point fallback for those two cells, pending user/template
  acceptance; the other three trait-34 grades retain 7 points. No new renderer
  dependency or PDF writer is introduced to hide that limitation.
- Access report font properties cannot be read/set during planning (2455).
  The renderer uses built-in Windows GDI printer metrics, verifies selected font
  faces and releases its owned fonts/DC. It no longer substitutes Courier for
  every field, estimates all fonts as 0.6-em, or relies on textbox reflow. Text is
  planned into individual lines, with grade wording flowing around the boxed marks.
  Continuations rewrap original text instead of retaining narrow field soft breaks.
  Chief dates follow verified YYMMMDD/MM/YY display
  pictures and SSNs use display separators; stored values remain unchanged.
- Existing nonblocking capture, pending-grid/editor precedence, alphabetical
  ordering, preview, combined/individual/Both output, signature exclusions,
  collision protection and transactional publication behavior remain in effect.
  Physical printed fit, an Acrobat-absent workstation and exact half-point style
  acceptance remain open. Earlier historical-output waivers are unchanged.

Development verification on 23 October uses the fresh
[0.36.0-dev/schema-33 blank](../build/dev/Swift_NAVFIT_0.36.0-dev_20261023-125710_96cbccd0.accdb),
SHA256 `C0E952EE5E648941E36A96AD24459246D3B2684296CEED96EE54122A428677E8`:

- [BuildAndTest receipt](../build/workflow/20261023-125710-45f282c4-e665-4f46-a79e-0c42c6c753c3/verification.json):
  compile/reopen, blank identity, exact source/bindings and standard native checks
  passed. The unchanged manifest has 212 inputs, 105 source modules plus five
  generated modules, 74 stages and 68 forms. Static catalogue/layout checks: 801.
- [Fresh-artifact PDF receipt](../build/tests/pdf-output-fb5de923-a8b0-4f32-99d2-fcf14e4208ea/verification.json):
  58 existing and 225 Chief-specific assertions passed without source installation.
  Independent PDF.js extraction checked 28 generated PDFs/390 Letter pages with no
  printed `#Error`. All 35 Chief grade explanations match the complete catalogue
  text, with no glyph overlap between lines or into numeric boxes. Blank, short,
  two-line and all-grade examples stay on two pages; the deliberate overflow case
  has six pages, all 10,500 oversized-comment characters and all 21 end markers.
  The 23-page combined matrix packet has its
  summary first and reports alphabetically ordered. Screen review covered Chief
  front/back pages, the densest grade-1 paragraph, the summary, and EVAL/FITREP
  front pages; physical output remains unverified.
- [Transfer receipt](../build/tests/major-69ee124b-dab5-41a5-94a5-5c72f602f7bc/verification.json):
  populated schemas 31, 32 and 33 import through Settings & Help into schema 33,
  preserving original values/history, relationships and native provenance;
  rollback, reopen, repeated-import and unsupported-source rejection passed.
  Chief-specific assertions: 107 for current schema and 129 for each older schema.
  Cold-recovery checks restored all 21 comments in a new Access session alongside
  invalid pending data.
- [Interoperability receipt](../build/workflow/h01-final-10d689e6-6104-4d8b-b418-686ef268d963/verification.json):
  298 structured-PDF checks, 549 native adapter checks, 410 exchange checks,
  121 integrity guards and 1,045 cumulative group/nonblocking-export checks passed.
  These prove native-format behavior, not a new NAVFIT98A desktop journey.
  Supplied build hashes remained unchanged. No runtime dependency, new runner,
  manifest stage, release publication or user-workspace migration was introduced.

### Remote regression and accessibility preparation

- In 0.41/schema 33, structured-PDF import and NAVFIT return services expose an
  optional typed outcome without changing existing callers. Their UI adapters
  distinguish Not started, confirmed rollback, Committed and uncertain outcomes.
  A post-commit/final-display failure preserves saved data, disables Apply, clears
  live ownership and gives close/reopen guidance; it never claims nothing applied.
  A precommit rollback retains reviewed choices and permits deliberate Resume/retry.
  A correctable return precondition retains the existing conflict-resolution flow.
- Members contracts require exact pending-cell retry/discard and export/PDF actions.
  My Work requires selection/open/draft/refresh and accept/return/cancel bindings,
  contributor and history/status controls. Queue loading asserts the contract before
  accessing controls, replacing obsolete silent probes for now-required controls.
  Build verification consumes the same literal contract data. No layout, schema,
  catalogue or runtime dependency is changed.

- In 0.40, workspace-import feedback uses a scalar snapshot of the diagnostic
  context instead of parsing display text. The service preserves the primary
  migration error number/source and passes its transaction outcome to the UI.
  Only confirmed rollback is reported as rolled back. A committed import with
  failed finalization retains data, clears live Apply ownership, disables retry
  and instructs close/reopen; unknown rollback outcomes also disable Apply.
  A new validation failure starts with its own outcome rather than inheriting
  stale committed diagnostics. Source reads remain read-only and verified.
- Four wizard contracts now require progress/status controls and exact critical
  actions: workspace import, structured-PDF import, NAVFIT return/group export
  and PDF output. Existing native/build verifiers consume the same literal
  contract data. Missing progress/status and misbound actions fail explicitly;
  form geometry, output serialization and schema remain unchanged.

- Reliability follow-up fixes the backup worker's obsolete schema-25 ceiling:
  embedded workers receive the application's supported ceiling; the standalone
  worker reads the canonical schema. ACCDB/ACCDE copies retain their extensions,
  verify hashes and all local table counts through read-only DAO, and reject
  linked tables, unsupported schemas and existing destinations/evidence. Primary
  failures survive cleanup; failure reporting never replaces existing evidence.
- PDF/native-return cancellation passes the live preview token to the existing
  contributor-owned deletion guard. Stale contributor and superseded same-owner
  contexts reject cancellation without deleting saved choices. Explicit Resume
  followed by Cancel still works. No schema, UI geometry or dependency changes.
- Recovery tests now make the mandatory fresh-session contributor choice through
  its actual callback. Read-row tests preserve deliberate corrupt-cache probes,
  assert membership rejection before any write, and restore their synthetic
  values before unrelated group mutations. No production integrity guard is relaxed.
- The existing Reports/My Work scale fixture likewise seeds traits through normal
  saves instead of diverging the revision/cache pair. Its optional maximum is
  100-1000 reports, retaining 25/100-row cases and five alternating-order samples.
  Per-run timings and verification evidence are unique; historical CSV files remain.
- An earlier refresh-ownership check failed on both unchanged 0.37 and 0.38:
  `UXWorkBack` returning to a closed My Work window produces three queue refreshes
  where the retained test expects two. In 0.39, a newly opened window uses its
  existing initialization/activation refreshes; an already-open window retains
  its explicit refresh. All seven ownership assertions remain intact. The 0.38
  receipt retains its historical failure; it must not be relabelled as passing.
- Workspace-import cancellation now passes its live token to the contributor
  guard. Reopening clears only live ownership and disables Apply until deliberate
  Resume. Fresh preview opens the form before staging its context, so initial
  OnLoad cannot clear that newly reviewed context. All three cancel adapters give
  safe retry/Resume feedback while preserving rejected stale choices.
- Native report/group export and retained-source rebuild use existing primary/
  cleanup diagnostics. Error number/source, redacted stages and owned pending
  paths survive journal/reader cleanup failures. File-publication state is explicit,
  including failure after files were published; retained output is never labelled
  absent or overwritten by retry. Report revisions are unchanged. Native fault
  tests combine missing templates with rejected failure journals and separately
  reject completion journals after publication, for both report and group output.

- The validation runner now establishes a native client size before asserting
  layout; this host can initially report zero width. The production pane also
  fixes an actual 200-twip overlap between evidence and its action buttons, sizes
  Help with the heading, and retains findings/evidence/action tab order. Findings
  and evidence have attached labels and access keys. Native resize/focus checks
  cover repeated narrow/wide sizes, selection and paging without data writes.
- Setup, contributor selection and PDF preview associate ten controls with their
  labels, including a new read-only PDF findings label. Report geometry remains
  unchanged. All 21 Chief comment controls retain their existing ordered tabs,
  attached labels and native focus behavior; no unnecessary tab-order rewrite
  was made. These are structural/focus checks, not screen-reader certification.
- Six permutations of PDF preview, saved NAVFIT export and findings reproduced
  an export-entry failure with invalid queued cells. Export now bypasses the
  editing tools' drain/leave-editor steps, preserving queued cells, raw editor
  values and recovery while opening the explicitly selected saved group. Senior/
  editing-page navigation still rejects invalid pending cells. Backend integrity,
  serialization, history, source and destination guards remain unchanged.
- A canonical Title proposal previously caused DAO 3265 when the trait cache
  tried to materialize a nonexistent Title column. Only matching cache columns
  are now projected. All proposals remain in the proposal list/detail and normal
  Apply path; preview does not commit them. Live/historical tests cover all three
  report kinds and stale-proposal rejection.
- Stale tests now use canonical Members navigation and an explicitly hidden
  workbench adapter. List headings are excluded from data-row counts, a Draft
  row is chosen after the earlier approval scenarios, and bounds are evaluated
  inside native VBA rather than through nullable external COM control properties.
  Ordered reviewers, the full 5,000-character staged memo, guards and original
  workflow assertions remain covered.
- The Settings workspace picker offers ACCDB and ACCDE. Both continue through
  the same read-only DAO preview/import handlers; source code is never executed.
  Compiled PDF automation remains unverified: bounded popup-control, COM argument
  and expression-engine attempts on the retained compiled fixture failed. The
  experimental driver was removed; no unreliable check, injected compiled VBA,
  automation framework or security-policy change was retained.
- Access Explorer inspected installed controls/tab metadata and contrast on a
  disposable copy. Its normal-view screenshot operation explicitly refused the
  validation PopUp before opening it; the owner closed normally. No alternate
  capture mechanism was attempted. Runtime visual, physical display/DPI,
  assistive-technology, normal-security and signing acceptance remain open.

Verified development artifacts and scope:

- [Fresh 0.44 ACCDB](../build/dev/Swift_NAVFIT_0.44.0-dev_20261024-121339_7f68a813.accdb),
  SHA256 `6E4FD9A69229EFE70084F86AB60ADE4A039E75B68D3ED2F6E1A898E6606EAFFB`.
  [BuildAndTest](../build/workflow/20261024-121339-8462f942-de34-4b86-8a0f-cc33ca379620/verification.json)
  passes fresh generation, compile/reopen, exact source/bindings, blank verification
  and all standard native suites. Schema 33/catalogue 2 and report geometry remain.
  The [final directory/list receipt](../build/workflow/directory-list-final-81ce79d7-826c-476f-9f61-1659209d1299/verification.json)
  passes 29 normal directory-batch assertions, isolated rollback/postcommit/display/
  simulated-uncertain/reset cases and reopen counts, 42 lifecycle assertions,
  28 group move/status assertions and all six reordered group-task scenarios.
  The [exact-artifact benchmark](../build/tests/summary-list-bd1b6131-d816-488c-a728-b3fbc87ced86/verification.json)
  uses production-verified fixture reuse and five alternating samples. All/Active/
  Archived medians were 347.72/88.00/50.02 ms on 0.43 versus 294.53/64.84/47.31 ms
  on 0.44. Workflow calls were 96/80/16 versus zero; displayed cells/order match
  and reads add no data/history rows. Parent artifact and fixture hashes remain
  unchanged. Counter/fault instrumentation exists only in labelled test copies.
- The 0.44 focused summary-list
  [five-trial receipt](../build/tests/summary-list-f89c9b22-d2d8-4e08-b4f5-a7c7dc5eec3a/verification.json)
  compares 400 reports/40 populated groups plus eight boundary groups. Every
  displayed cell/order matches for All/Active/Archived and data/history counts
  remain unchanged. All-filter workflow lookups fall from 96 to zero; five-sample
  median All/Active refreshes were 309.48/73.62 ms versus 271.72/59.62 ms.
  The [fixture-reuse receipt](../build/tests/summary-list-80cb5e2d-d2af-4503-866b-2688802e1acd/verification.json)
  additionally verifies complete production migration/readback and unchanged
  fixture hashes. These are local measurements on refreshed copies, not a
  performance guarantee or final-artifact acceptance.
  The [directory fault/reset receipt](../build/tests/group-correctness-5ae429bd-e49f-4925-87c4-fd599a53abc2.accdb.cases/verification.json)
  covers real rollback, postcommit/display failure, simulated uncertainty, a real
  reset, duplicate-Apply rejection and close/reopen report/link counts. Normal
  directory tests pass 29 assertions and lifecycle tests pass 42. One earlier
  reused-controller fault run crashed Access with native exception `0xc0000005`
  in ntdll; its primary RPC and separate cleanup failure remain recorded. A new
  owned controller per isolated case passed all cases. Root cause is not claimed.
- [Fresh 0.43 ACCDB](../build/dev/Swift_NAVFIT_0.43.0-dev_20261024-112635_4ce59eed.accdb),
  SHA256 `C1DCB575F48DD0431D1A03FC1159E9E22DFA1391F939B43F74983BA078582089`.
  [BuildAndTest](../build/workflow/20261024-112635-e94c5237-87a6-47d5-a9df-a707c8e2d635/verification.json)
  passes fresh generation, compilation/reopen, exact source/bindings, blank
  verification and all standard suites, including the expanded 53-check
  directory-merge suite. Schema 33/catalogue 2 and report geometry are unchanged.
  No release, signing or physical acceptance is claimed.
  The [final preview receipt](../build/workflow/owned-preview-final-16236959-61e5-46b1-b7b1-9f290ed152e0/verification.json)
  repeats native group import, four outcome-fault variants, real three-preview
  reset/reopen checks, move ownership/reopen, committed move/merge display faults
  and all 55 diagnostic/view-contract assertions against this exact artifact.
  Normal cases use installed source; explicitly labelled fault copies receive
  synthetic instrumentation. The parent artifact hash remains unchanged.
- The [0.43 focused preview receipt](../build/workflow/owned-preview-focused-8ffa518a-d770-4a7a-94fb-5fd4d9d05e5b/verification.json)
  passes 274 native group selection/atomic-import/readback assertions, four
  import outcome-fault cases, a real reset with all three previews open,
  close/reopen report/group/receipt counts, 28 move/filter/status assertions,
  53 duplicate-merge/history assertions and 55 diagnostic/view-contract assertions
  plus installed rename/wrong-callback rejection. Cross-contributor Apply was
  reproduced on unchanged 0.42 before the fixes. Uncertain outcome is simulated;
  no actual failed DAO rollback is claimed. These are refreshed disposable-copy
  checks, separate from the final fresh integrated artifact.
  The [display-fault receipt](../build/workflow/preview-display-faults-537338ec-4f3f-451f-9751-3f4a84216798/verification.json)
  adds seven committed-move and eight committed-merge/display/Undo assertions.
  Faults exist only in the owned test copies, not maintained application code.
- [Fresh 0.42 ACCDB](../build/dev/Swift_NAVFIT_0.42.0-dev_20261024-105339_4e40404e.accdb),
  SHA256 `0DD14857BBEAC6201F9FD76AC271CCF5066F814FE512B83C6430A2CE1C1EE2C8`.
  [BuildAndTest](../build/workflow/20261024-105339-86f26ada-3e25-4931-8451-3f98297c5b13/verification.json)
  passes fresh generation, native compile/reopen, exact installed source and
  bindings, blank verification and all standard native suites. Schema 33 and
  catalogue 2 remain unchanged. The prior 0.41 artifact and tracked native blank
  retain their original hashes. This is a development artifact, not a release.
  The [exact-artifact focused receipt](../build/workflow/ui-feedback-final-69f4e3fa-79df-4b59-8a9b-7d92b4a23fb5/verification.json)
  passes 73 NAVFIT UI assertions, including internal-template verification,
  restricted visible controls, nonblocking export and summary-list PDF capture
  of the selected roster with combined-packet/summary defaults. It also repeats
  fixed pages/labels, real-reset recovery, group navigation/task orders and
  preserved folder-history/guarded-undo checks without installing application
  source. The final artifact hash is unchanged after testing.
- The [0.42 focused candidate receipt](../build/workflow/ui-feedback-focused-9e013edb-f460-457a-8170-dae13d9f24c5/verification.json)
  passes ten native runners: 407 fixed-page assertions, ten label/tab checks,
  15 real-reset trait/recovery/stale-baseline assertions, group/exit task orders,
  feedback navigation, preserved folder history/undo, nonblocking native group
  export, live Members, close/reopen recovery, audit/conversation navigation,
  and required-view diagnostics. The
  [conversation/close receipt](../build/workflow/ui-feedback-close-f8fb463c-5579-41c3-a517-ad0a3ba39480/verification.json)
  adds 47 report-tools assertions, 43 native-close/workflow assertions and
  forced-close/reopen preservation of all 21 Chief comments with invalid raw data.
  These checks used disposable candidates, not a published release.
  After visual review exposed late presentation decorators, the final export
  layout passed 69 native UI assertions on a focused layout copy. Access Explorer
  screenshots confirmed its static layout; that Explorer session had code
  disabled, so its screenshots do not establish interactive or Chief runtime
  visual acceptance. Native geometry/behavior checks are separate from physical
  display/DPI and printed-fit acceptance. No security settings were changed.
- [Fresh 0.41 ACCDB](../build/dev/Swift_NAVFIT_0.41.0-dev_20261024-005848_6426d8fa.accdb),
  SHA256 `9727B09EF26DD9D10952AB25B6991BDA9951B45A388EF820B4C4AD5ABEA1D00F`.
  [BuildAndTest](../build/workflow/20261024-005848-9fd9b88c-f27c-461c-8d36-57ce76fcc355/verification.json)
  passes all 74 stages, compile/reopen, installed source/bindings/stronger contracts,
  blank verification and standard native suites. Schema 33/catalogue 2 remain.
  The [apply-fault receipt](../build/tests/wizard-state-dc09054a-de74-44f7-af0d-6c4ba2b7badf/verification.json)
  passes 151 outcome/retry/persistence assertions across eight isolated PDF/return
  cases, including normal saved revision sequences, one receipt and stable counts
  after close/reopen. Uncertain outcomes are simulated; no actual failed DAO
  rollback is claimed. Postcommit and display faults are explicit local test-copy
  instrumentation, not shipped hooks.
- The [0.41 reset/restart receipt](../build/tests/wizard-state-bded9715-c6e8-4008-a104-3280f7633ab8/verification.json)
  passes 63 assertions and a real schema-29 import, with source hashes intact.
  Its intentional visible unhandled-error injection is explicitly skipped. The
  earlier focused default run required assisted End of the verified synthetic error,
  then completed reset/restart successfully; this is not unattended UI-error proof.
  Separate user Access workspaces were neither closed nor modified.
- The [final 0.41 focused receipt](../build/workflow/apply-contracts-final-db21289f-9ea7-42ca-84f0-9e7372873539/verification.json)
  passes all ten selected runners on this exact artifact without application source
  installation. It includes the 151 apply-fault and 63 reset/restart assertions,
  43 diagnostics/contracts plus both installed view faults, 1,748 My Work, seven
  refresh-owner, 73 live Members, 118 unified Members/exit and 162 reordered-task
  assertions, 549 adapter/410 exchange/147 guard checks, 48 NAVFIT UI checks and
  1,045 cumulative group/nonblocking export checks. The explicit test-copy fault
  instrumentation and refresh counter are recorded separately from production code.
  Source hashes remained unchanged. No dependencies, release, binary publication,
  physical print, signing/security or attended accessibility acceptance are added.

- [Fresh 0.40 ACCDB](../build/dev/Swift_NAVFIT_0.40.0-dev_20261023-233901_08795eef.accdb),
  SHA256 `9DB8F7B4C549BEFD6EC66C12F99A25DD6323E43920815FCBC942736E47932DE5`.
  [BuildAndTest](../build/workflow/20261023-233901-94949ce1-e0df-475b-8297-f56a3478994f/verification.json)
  passes all 74 stages, native compile/reopen, installed source and stronger view
  contracts, blank identity and standard native suites. Source remains schema
  33/catalogue 2. The focused wizard source installer now includes the diagnostic
  module with the updated importer, avoiding an undefined helper on partial test
  copies. Fresh generation installs and verifies the complete module set.
- The [final 0.40 receipt](../build/workflow/import-contracts-final-1a1cc3cc-57a8-4730-8670-9185a173c6be/verification.json)
  passes eight focused runners against that artifact without the partial-source
  install switches. The engine suite passes 16 adversarial comparisons, real
  precommit mismatch rollback/retry and 14 committed-outcome/primary-error/UI/
  source-protection assertions. Its post-commit fault is explicitly injected at
  the real production commit boundary on a separate disposable copy, not an app
  hook or a replacement migration implementation. The contract suite passes
  35 assertions plus both installed rename/wrong-callback faults; wizard tests
  pass 57 task-order and 63 reset/error/restart assertions. The
  [transfer matrix](../build/tests/audit-message-transfer-d72a45bc-a2f3-4376-ad12-996a4d9b8d78/verification.json)
  passes actual retained schemas 28-33 with original rows/fields, histories and
  source hashes intact. Native adapter/exchange/guards pass 549/410/147,
  group/nonblocking export 1,045 cumulative, and recovery 12+26 assertions.
  Existing historical-fixture waivers and physical/security/accessibility gaps
  remain; no new dependency, maintained script, release or binary upload is added.

- [Fresh 0.39 ACCDB](../build/dev/Swift_NAVFIT_0.39.0-dev_20261023-220519_90c5185e.accdb),
  SHA256 `BE47C13FEAEDC4C37CAFF233C1F461170149A2D65DCFC3276287BD7FD80494CC`.
  [BuildAndTest](../build/workflow/20261023-220519-8c7a47ff-64b1-4420-ade9-d08154894af9/verification.json)
  passes all 74 stages, native compile/reopen, exact source/bindings, blank identity
  and standard native suites, including fixed-page/pending/workflow checks.
  Source remains schema 33/catalogue 2, with no new runtime dependency or manifest
  stage. This is development evidence; printed fit, normal-security/signing,
  compiled-PDF and attended accessibility acceptance remain open.
- The [final 0.39 reliability receipt](../build/workflow/m02-m04-final-bb3f1e35-17fe-4454-865e-d6a74eb89add/verification.json)
  passes all 11 selected runners against this exact artifact: seven refresh-owner,
  1,748 My Work, 198 Reports, 12+26 recovery, 57 task-order, 63 reset/error/restart,
  549 adapter, 410 exchange, 147 export/fault guards, 1,045 cumulative group/export
  assertions, and 19 diagnostics/contract assertions plus both installed view faults.
  Existing harnesses and the explicit refresh-counter instrumentation were used;
  final checks did not replace application services. Native diagnostic tests prove
  primary/cleanup separation, late rollback/retry and post-publication recovery.
  The [transfer matrix](../build/tests/audit-message-transfer-606bb906-2c93-4467-82ae-b6d8e19534ea/verification.json)
  passes six actual retained schemas 28-33, with original fields, replies/private
  drafts, actor/event identities, idempotency, reopen and unchanged source hashes.
  The [backup receipt](../build/tests/backup-18d8376c-583b-4439-a727-16316d0fbe99/verification.json)
  passes 14 fault/round-trip cases and all 21 Chief recovery comments. Its compiled
  format fixture is the retained 0.38 ACCDE; this is not compiled 0.39 acceptance.
  No historical fixtures 1-27 were recreated. No owned Access session remained.

- [Fresh 0.38 ACCDB](../build/dev/Swift_NAVFIT_0.38.0-dev_20261023-202007_dd3b552e.accdb),
  SHA256 `7A1338659C05D62FF246B4EA22AF965617FEEBBE751942EF9AE12D88FA95D58C`.
  [BuildAndTest](../build/workflow/20261023-202007-a9c6b4d7-5499-4930-ae09-07c1a0dddb0c/verification.json)
  passed all 74 stages, compile/reopen, exact installed source/bindings, blank
  verification and the standard native suites. No build inputs changed afterward.
- The [0.38 reliability receipt](../build/workflow/reliability-final-7a4a1349-6323-4167-9743-a95d3a2b2187/verification.json)
  retains recovery 12+26, Reports 198, My Work 1,748, state/performance invariants
  16, task-order 36 and full wizard reset/error/restart 63 passing assertions.
  Application source was not installed over the final artifact. The separately
  reproduced refresh-ownership failure above is explicitly retained, not a PASS.
- The [1,000-report scale receipt](../build/tests/read-scale-c516b461-e2d4-4860-9851-233604285165.accdb.verification.json)
  records exact prepared/unprepared values and ordering and no read-history writes
  at 25, 100 and 1,000 reports. Five-sample medians at 1,000 were Reports
  546.9 ms prepared / 23,351.6 ms unprepared, and My Work 335.9 / 25,179.7 ms.
  These compare existing read paths on this workstation, not new optimization
  gains, screen-paint timings or tail-latency guarantees.
- The [populated major transfer receipt](../build/tests/major-cecd12d2-5967-40ce-b659-11d1255fd1e8/verification.json)
  passes current/schema-31 source workflows, Chief comments/native provenance,
  exact original fields, complete relationship signatures, late rollback, reopen,
  populated-target/future-schema/unknown-table rejection and source hashes.
  The [audit/message matrix](../build/tests/audit-message-transfer-ffa6f99f-1f66-4544-9287-71c27d8cb757/verification.json)
  passes actual retained schemas 28-33 with reply/private-draft/actor/event
  identities and idempotent reopen; historical outputs 1-27 were not recreated
  or claimed retested. No release or attended acceptance is implied.
- The [0.38 backup receipt](../build/tests/backup-74db632c-3d9f-411c-8618-fc439e0f0d6f/verification.json)
  passes 14 close/copy/restore/fault cases against the final artifact without
  source installation, including all 21 Chief recovery comments, raw dates,
  literal paths, linked/corrupt/future-schema rejection, protected existing data/
  evidence and byte-identical ACCDB/ACCDE round trips. Original hashes remain intact.
- [Unsigned 0.38 x64 ACCDE](../build/dev/Swift_NAVFIT_0.38.0-dev_x64_20261023-204851_6adab3dc.accde),
  SHA256 `664F2EE31EA00DD8A0A1F8CB85E35B5757541AE3D123511E71B6A1810A813359`.
  The [Compile receipt](../build/workflow/20261023-204811-6c9774a4-b992-4874-aea2-7a4298f169dd/verification.json)
  and [compiled checks](../build/tests/compiled-f14b593f-263f-421e-943b-0ada1617396d/verification.json)
  pass 34 trusted desktop-native assertions on Access 16.0.20430.20182 x64.
  Signing, normal-security and compiled-PDF acceptance remain open, as do existing
  physical/accessibility gates. No new script, framework, runtime dependency,
  manifest stage or release publication is included in that batch. A later GitHub
  source checkpoint does not publish these local binaries.

- [Fresh 0.37 ACCDB](../build/dev/Swift_NAVFIT_0.37.0-dev_20261023-182934_03afa372.accdb),
  SHA256 `068ED131B3B01AC9CAE598FD6691FC9182DBD4802845CBF5E4CE26F34E8AC048`.
  [BuildAndTest](../build/workflow/20261023-182934-e8dc60a9-c0bd-4475-83a7-e9c74e206536/verification.json)
  passed compile/reopen, exact source/bindings, blank identity and standard native
  suites, including 354 fixed-page/focus checks and ten entry-label associations.
- [Final regression receipt](../build/workflow/remote-regression-final-c7a1d845-aed9-45ab-be6b-4d10ebe15272/verification.json)
  checks that artifact without source installation: validation 141; retained
  adapter six plus 92 native bounds; live/historical Members 73; unified Members
  118 cumulative plus 162 six-order checks; approved-trait retry 26; NAVFIT UI 48;
  native group/nonblocking export 1,045 cumulative; PDF 58 existing and 225 Chief
  assertions plus the print matrix. Independent extraction checked 28 PDFs/390
  Letter pages, expected Chief page counts and all 10,500 overflow characters.
- [Unsigned 0.37 x64 ACCDE](../build/dev/Swift_NAVFIT_0.37.0-dev_x64_20261023-184717_142f8627.accde),
  SHA256 `306ED2D901E2EB27B18CE0532791C2029E41A8070CC9BE5C28E9CC9FEE432328`.
  The [Compile receipt](../build/workflow/20261023-184637-efe996a2-6ed6-40de-be14-f4c964471675/verification.json)
  and [compiled checks](../build/tests/compiled-2d597ba5-fbf2-4068-a52e-a0c8efed9dc7/verification.json)
  record 34 trusted native assertions on desktop Access 16.0.20430.20182 x64,
  including ACCDB/ACCDE import, save/reopen, source protection and provenance.
  They explicitly record normal-security and compiled-PDF testing as false.
- [Final installed-object audit](../build/tests/m07-final-37db6ff2-8381-4481-a682-cc210178266c.accdb.audit.json)
  reports no static heuristic contrast/truncation findings on setup, contributor
  selection, PDF preview or validation. This is not runtime visual acceptance.
  Existing schema/physical-test waivers remain; no release, signature, upload,
  framework, dependency or new persistent runner is introduced.

### L01–L02 relationship names and revision-ID aliases

- Schema 32 replaces the 34 opaque major relationship names with
  `FK_<child table>_<foreign field>`. Referenced tables/fields, field order,
  enforcement and cascade attributes remain unchanged. Upgrade uses the existing
  Settings & Help import into a fresh blank; the source file is read through
  read-only DAO and its schema/code is not modified or executed by import.
- [major-schema.json](../src/major-schema.json) defines two simple, writable saved
  queries: `qryChangeItemRevisionIDs.ExpectedRevisionID` aliases the original
  `ChangeItem.ExpectedRevision`, and `qryMeetingProposalRevisionIDs.BaseRevisionID`
  aliases `MeetingProposal.BaseRevision`. These are `ReportRevision.RevisionID`
  references, not `RevisionNumber`. There is one stored value, with no duplicate
  columns, synchronization writes, renumbering or backfill. Legacy physical
  columns/FKs remain available; production import copies and verifies them exactly.
- Batch preview/apply/undo, meeting proposal/apply, group-impact and comparison/
  workbench queries consume the aliases. Existing stale-revision, ownership,
  atomic rollback and history behavior remains in effect. Historical source
  queries are not trusted/copied; target queries come from the reviewed schema.
- The existing builder installs the schema-owned query definitions. Reopened
  installed verification checks relationship definitions/enforcement, absence of
  opaque names, and exact alias SQL. No manifest stage or dependency is added.
- The existing major runner checks bidirectional alias writes, FK rejection,
  rollback and IDs distinct from revision numbers, alongside existing workflows.
  Optional schema-31/30 inputs exercise populated Settings imports, exact original
  field verification, complete DAO relationship-signature equivalence, late
  rollback, populated-target rejection, reopen, future-version/unknown-table
  rejection and input hashes. Old-output testing waivers remain unchanged.
- The current comparison preview/detail and retained hidden workbench adapter are
  exercised directly. The old visible-workbench UI fixture failed identically on
  the unchanged 0.34 baseline at its report-selector assertion; 0.37 subsequently
  repairs its hidden-context/header assumptions while retaining its assertions.
  An initial
  alias harness sequence triggered native DAO 3246; owned DAO readback and
  rollback-before-cursor-close retain all alias/rollback assertions without a
  production workaround. Only identified synthetic test sessions were recovered.
  The older live-members fixture also failed with missing-workbench error 2450
  on both the candidate and unchanged 0.34 baseline. Its 0.37 replacement uses
  canonical Members navigation and retains all live/historical/guard checks.
  Earlier comparison proof used Trait33; a generic Title proposal hit error 3265
  because the trait-cache column set does
  not contain Title. It was outside L01–L02 and was not counted as passing at
  that checkpoint; 0.37 repairs it as described above.

The [fresh 0.35.0-dev/schema-32 blank](../build/dev/Swift_NAVFIT_0.35.0-dev_20261023-100937_1dafb477.accdb)
has SHA256 `2B2D41919E6DA99937BE9EC739DDDE69031CD5B138369A7391172545D89BC21A`.
Its [integrated receipt](../build/workflow/20261023-100937-25efe742-9f94-4371-add8-427a18aa7a1d/verification.json)
passes all 74 stages, compile/reopen, installed source/bindings, blank checks,
108 named single-column relationships and both canonical alias queries; standard
report-tools (43), contributor (61), fixed-page (264), workflow/recovery (41),
directory-merge (47), smoke, reopen/import/cancellation checks also pass.
The [exact-artifact major receipt](../build/tests/major-bd224c7f-f4ee-4a88-b5fd-8e5641d8e5b7/verification.json)
records 46 existing workflow, 33 alias/rollback/FK and 13 comparison/adapter
assertions, plus populated schema-30/31/32 Settings imports, unchanged original
fields/IDs and complete relationship signatures, late rollback, reopen,
populated-target/future-version/unknown-table rejection and source hashes.
No application source was replaced. Final continuation checks also pass all
26 approved-trait retry assertions and the
[NAVFIT adapter/exchange/guard suite](../build/tests/navfit-adapter-b4150473-825a-4306-9395-7ab225bcd244/verification.json)
(507/410/100 assertions). No release or physical/printed acceptance is claimed.

### H02 nonblocking NAVFIT98A export

- Report and group exports include saved revisions despite incomplete fields,
  mismatched group keys/type, readiness findings or archived/approved status.
  The exchange selectors include archived records. Findings remain informational,
  with existing report/group issue views available; no acknowledgement is added.
- Every saved group member is exported using its own form kind and values.
  Missing recommendations remain null, are disclosed separately and contribute
  to no recommendation count (including NOB). An unavailable/partial/estimated
  PSG RSCA is disclosed in preview and preserves native RSCA instead of blocking
  export or inventing an average; the frozen baseline records actual PSG use.
- Saved-only service exports do not touch open editor values or recovery drafts;
  screen navigation retains explicit save/discard/cancel. Missing report end
  dates use the existing `UNSPECIFIED` filename convention. Export does not
  constitute approval, signature or submission.
- Native format capacity/type constraints, sealed provenance, stale revision/group
  checks, collisions, destination access and truthful failure remain enforced.
  Unsupported checkbox/enumeration/date values cannot be silently coerced.
  Group-return reconciliation remains unsupported. No schema, dependency,
  manifest or release change is included.
- Existing NAVFIT adapter/round-trip, group-export and exchange-screen runners
  cover complete serializable fixtures and incomplete/mixed archived groups,
  approved reports, exact/unavailable PSG, null values, pending buffers/recovery,
  source/history preservation, stale previews and failed destinations. Their
  default checks installed application code; `-InstallSource` is only for early
  disposable-copy development checks.

The [fresh 0.34.0-dev/schema-31 blank](../build/dev/Swift_NAVFIT_0.34.0-dev_20261023-091723_cb948079.accdb)
has SHA256 `E151101B08C5CC81FA1ADC276A7D9595784A9F552593AD5A130379A07C6F8353`.
The [integrated receipt](../build/workflow/20261023-091723-b304841e-81b8-4d51-b144-ae11dba0215d/verification.json)
passes all 74 build stages, compile/reopen and current-source/blank verification,
43 report-tools, 61 contributor, 264 fixed-page, 41 workflow/recovery and 47
directory-merge assertions plus smoke/reopen/import/cancellation checks.
Exact-artifact checks, without application-source replacement, pass:
[group export](../build/tests/navfit-group-e64587c5-b6d3-4dd0-8c5c-c93b8a0f9510/verification.json)
(982 cumulative assertions plus 507 adapter assertions),
[exchange screen](../build/tests/navfit-ui-e7125b63-15ad-4738-9a5f-41031ecb11bc/verification.json)
(48 assertions), and
[adapter/round-trip/guards](../build/tests/navfit-adapter-0bfaa07b-609f-4d63-a7a3-2f8d62868bfd/verification.json)
(507/410/100 assertions). The existing filename suite passes all nine checks.
The older guard harness expected pre-M02 error text; baseline reproduction and
inspection confirmed the durable wizard rejects changed source/target first.
Assertions now expect those specific errors and additionally verify the rejected
return retained the current revision/comments. No production guard was weakened.
Native screen behavior was exercised; no new screenshot/layout acceptance,
NAVFIT98A desktop journey, printed-fit or release acceptance is claimed.

### M04–M06 diagnostics, view contracts and runtime trait data

- M04 retains existing numeric error codes and adds module/procedure identities
  to explicit application raises. Transaction handlers capture the original error
  before cleanup, record stage and rollback outcome, and retain secondary close,
  rollback and failed-journal errors separately. The owned diagnostic context is
  in memory; surfaced diagnostics and failed-operation outcomes redact compact,
  hyphenated and space-separated personal identifiers. No raw diagnostic log or
  new logging table is introduced. Successful saves and retained audit data are
  unchanged. Editor save failures distinguish committed writes from refresh errors.
- Cleanup helpers catch their own failures. Redundant broad suppression around
  those calls was removed; remaining `Resume Next` sites are bounded optional
  focus/text/property probes or local cleanup/recovery. The focused runner emits
  complete suppression and reused-code inventories for review. Codes were not
  renumbered, and Boolean event callbacks remain supported.
- M05 adds explicit required control/event contracts for the editor, Members
  grid, import previews and PDF output. Canonical field metadata owns editor
  control membership; trait-grid contracts verify queue callbacks. Optional
  compatibility probes remain distinct. Build verification checks installed
  modules, literal form/control references, form/report expressions and macro
  callbacks. Missing/renamed controls and wrong existing callbacks fail clearly;
  literal contract extraction and native-property inspection avoid executing VBA
  during build checks or changing file-trust policy. Editor/grid entry checks
  protect the pending-buffer boundary without changing geometry or navigation.
- M06 adds `TraitDefinition` with a composite form/block/mark/catalogue-version
  identity, constrained domains, exact title/description, source identity and
  original trait ID. The reviewed JSON seeds 147 reference rows (21 traits ×
  general guidance plus six marks); `GENERAL` is an explicit storage key, while
  callers retain the empty-mark API. NOB descriptions remain Null. Runtime readers
  use the table; the generated module now contains only identity and thin accessors.
  Official strings remain exact; editor display removes embedded title line breaks
  as before. This does not add the missing authored Chief performance comments.
- Blank means no user/domain data: the blank verifier permits only the exact
  reviewed 147 reference rows. Production import retains the current seed for
  older schemas, preserves additional source catalogue versions, verifies copied
  reference values before commit, and rejects malformed source column types and
  same-key conflicts atomically. Lookup numeric keys are explicitly typed.
  Unknown source tables and populated targets remain rejected; historical import
  support and the old-output waiver are unchanged. No manifest stages, runtime
  dependencies or release artifacts are added.

The fresh 0.33.0-dev/schema-31 artifact is
[the verified development database](../build/dev/Swift_NAVFIT_0.33.0-dev_20261023-081452_45345334.accdb),
SHA256 `B1FAF44700EF85BD56C3BDA0DC262F637467498F69A329F3B870B47613EA911C`.
Its [build receipt](../build/dev/Swift_NAVFIT_0.33.0-dev_20261023-081452_45345334.accdb.build.json)
records all 74 stages, compile/reopen, 105 static modules and 68 forms. The
[standard development test receipt](../build/workflow/20261023-084219-063e9c12-8fe0-48cb-8c9a-8ae962fe62d5/verification.json)
passes blank/reference verification, 43 report-tools, 61 contributor, 264 fixed-page,
41 workflow/recovery and 47 directory-merge assertions plus persistence smoke checks.
Exact-artifact [diagnostics/contracts](../build/tests/diagnostics-contracts-aaa02d91-5076-4d16-a669-9d3702b485d4/verification.json)
pass 19 behavior assertions and both installed-verifier negative checks; the
[catalogue/transfer receipt](../build/tests/trait-transfer-069ea664-f07e-4c44-bdc7-d5cccf3ede89/verification.json)
passes 909 exact/reference checks, populated schema-30/31 transfers, retained
versions, same-key conflict rollback, malformed-column rejection and source hashes.
No application source was replaced for these exact-artifact checks.
The [final continuation receipt](../build/workflow/m04-m06-final-continuation-20261023-085103/verification.json)
also passes 26 trait-retry assertions and 58 PDF-output assertions, including
pending drafts, failed destinations, native preview and the 144-member summary.
Those PDF checks do not close H01's known Chief-comment/print-fidelity defects.

One fixed-page run reported DAO 3059 with rollback; the unchanged baseline,
diagnostic copy and subsequent unchanged-artifact runs passed all 264 assertions.
One trait-retry run reported an open failure before its first assertion; subsequent
unchanged-artifact runs passed all 26. Root causes of these transient observations
are not established; failure logs remain retained and no assertion was weakened.
Synthetic demo/provenance fixture guards now distinguish schema-31 reference rows
from user data; normal blank verification still requires the exact reviewed seed.
Earlier schema-fixture gaps/waivers, physical display/printing, accessibility and
release acceptance remain separate. No release or publication is implied.

Desktop Windows, Microsoft Access, Windows PowerShell 5.1 and Git are the current
development environment. No free Access Runtime, macOS/WSL application runtime,
production ACCDE distribution or signing acceptance is established. Python/Node are tooling
dependencies for particular helpers, not requirements of native PDF extraction.
Access Explorer and `swift_files` are already available; no new tooling is needed
for the consolidation. Use disposable copies for native investigation.

M01 adds an optional `Compile` action to [develop.ps1](../scripts/develop.ps1),
converting a disposable copy of a verified current blank ACCDB into an unsigned
ACCDE development candidate. It leaves authoritative generation and the ACCDB
build contract intact. [test-compiled-distribution.ps1](../scripts/test-compiled-distribution.ps1)
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
The workspace picker now lists ACCDB and ACCDE in 0.37; direct production handlers
continue to import ACCDE data. Existing ACCDB Blank/Demo publication, historical import and release gates
remain required. [The workflow](../docs/Development_Workflow.md#compiled-desktop-candidate-m01)
records compilation, signing/package integrity and remaining acceptance steps.
The [compiled development receipt](../build/workflow/20261022-170507-27d28ca5-1cff-419a-ada3-49047c5dd297/verification.json)
records 34 passing trusted native assertions and the accepted candidate hash.
The overwrite guard also passed; the supplied ACCDB hash remained unchanged.

[develop.ps1](../scripts/develop.ps1) is the standard fresh build/test entry point;
[tasks.json](../.vscode/tasks.json) makes it the default build task. It also supports
Preflight, Build and Test actions. [release-check.ps1](../scripts/release-check.ps1)
checks supplied Blank/Demo artifacts; it neither creates the Demo nor publishes.
[build-ux-admin-demo.ps1](../scripts/build-ux-admin-demo.ps1) is the single supported
synthetic Demo builder: verified blank in, new seeded copy out, followed by
populated-Demo and startup checks. See the [Demo guide](UX_DEMO_BUILD_GUIDE.md).
The older preview builder is quarantined; its distinct feature suites remain
available for change-specific verification instead of running on every seed.
See [Development_Workflow.md](../docs/Development_Workflow.md) for current commands.

Artifact verification now reuses the builder's existing structural receipt only
for an unchanged artifact and complete matching input/checker fingerprint.
Metadata and requested blank inventories always run; workflow and migration
checks remain separate. Missing/stale receipts trigger native verification, and
`-ForceNative` requests it explicitly. No additional cache/dependency was added.

Deliver newly named Blank and synthetic Demo under [output/](../output/), preserve
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
[CONTRIBUTOR_STARTUP_R128_VERIFICATION.json](CONTRIBUTOR_STARTUP_R128_VERIFICATION.json).
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
Verification on the fresh source build: [development workflow receipt](../build/workflow/20261022-002149-1abcfadc-3df4-4471-9704-7e5d6362fd86/verification.json)
passes build/compile/reopen, installed source and bindings across 67 forms, blank
inventory and the standard workflow suites. The expanded
[readiness harness](../tests/ReadinessHarness.bas) passes 150 assertions on a
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
The [final development receipt](../build/workflow/20261022-015423-b7d4be0c-ca65-4f69-b030-55257beeca5f/verification.json)
passes fresh build/compile/reopen, installed source/bindings across 102 modules
and 67 forms, blank inventory and standard workflow suites. The final blank's
SHA-256 is `72CF2A7E011D36797E114799B7A20BA1F546D26505405BC731BE745D4D0173FC`.
The [focused logs](../build/workflow/c05-c07-20261022-020443-3da1242d-7b0d-4f80-b3fe-9c77c7108690/)
record 34 persistence, 43 accounting, 569 catalogue, 200 default, 507 NAVFIT
adapter, 410 exchange, 99 interchange-guard, 12 PDF counseling and 16 engine
assertions, plus populated schema-27/28 transfer and late rollback.
The [boundary receipt](../build/workflow/c05-c07-final-20261022-022006-63fe221a-e2cb-4b97-8b99-54815e9d6660/verification.json)
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
[producer/consumer map](CONSOLIDATION_PLAN.md#c08c10-producer-consumer-and-navigation-audit)
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
[audit/conversation/navigation](../scripts/test-audit-conversations-navigation.ps1)
and [retained-schema transfer](../scripts/test-audit-message-transfer.ps1).
The NAVFIT UI runner now points to the retained reference template rather than
its obsolete root-folder path; the original harness assertions are unchanged.
The [final fresh build receipt](../build/workflow/20261022-081737-c354c3e3-1221-43ea-845a-36ed661924d3/verification.json)
passes all 72 stages, compile/reopen, installed source/bindings (103 static
modules, 67 forms), blank inventory and the standard development suite.
SHA-256: `F8B6883CCFE9DEABF9DE2C811DD19DFEACDD2C6F54B91C5B92B946C5E93994B9`.
The [focused receipt](../build/workflow/c08-c10-final-20261022-0827/verification.json)
links 1,988 audit/conversation/navigation, contextual-help, My Work, group and
cleanup assertions on that exact artifact, plus the native form/macro inventory.
The [transfer matrix](../build/tests/audit-message-transfer-fa66cd5c-a0e3-4544-8acc-49d7218a7502/verification.json)
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

## Codex review fixes R01–R08

Fixed on 1 November 2026, from the Codex review of 0.51.0 (see the
[backlog](PRIORITY_BACKLOG.md)):

- **R01, status precedence.** The batched Reports reader no longer lets a review
  overlay replace Signed or "Reopened for correction", and it shows a report whose
  signoffs are all reopened as "Reopened for correction", matching
  `UXReportState`. Review assignments and routes on a signed report leave My Work
  (`UXWorkVisible` and `UXWorkVisibleRows`) and return when it is reopened. The
  unused participants query in `UXWorkVisibleRows` was removed.
  `UXReadRowsHarness` covers signed, reopened and re-signed reports with a
  pending assignment and route.
- **R02, completed groups.** `AddGroupMember` and `RemoveGroupMember` refuse a
  Completed group ("Reopen the completed group before changing its members").
  Completing a Completed group again with a signed date signs any member still
  unsigned. `SignoffHarness` covers both.
- **R03, backup journal failure.** When the journal entry fails after a backup
  and its verification file are written, `WriteWorkspaceBackup` keeps both and
  reports that only the journal entry failed. An earlier failure removes the
  backup and any verification file this call created. `test-recovery` injects
  the journal failure and checks the kept file against its receipt hash.

- **R04, prior evaluation.** An archived Swift report counts as prior history
  only if it was ever marked signed; an abandoned draft no longer supplies the
  prior ITA. Both `PreviousMemberEvaluation` and its batched twin apply the rule.
- **R05, RUAD failure diagnostics.** `RUADApply` raises the primary failure
  (for example DAO 3317) instead of error 450.
- **R06, RUAD preview owner.** A roster preview applies only for the contributor
  who made it; after a contributor switch Apply asks for a new preview.
- **R07, Duplicate.** A copy keeps the original's member link and match state
  (`DirectoryCreationContext` now carries both), so two members with the same
  name no longer produce a third profile.
- **R08, list punctuation.** The shared `UXListCell` quotes value-list cells, so
  commas, semicolons and quotes stay in their column in the NAVFIT group and
  report pickers and the NavyWriter list. The PDF output list uses it too.
- **T01.** `test-ruad-import` generates a synthetic 35-column NRRM roster when no
  `-SamplePath` is given, so it runs in every external-import suite run.

R01–R03 verified on build `Swift_NAVFIT_0.51.0_20261101-195039_c961a8ad`: BuildAndTest
PASS, plus recovery, summary-groups, workspace-read-performance,
workspace-resources and reporting-senior-accounting suites, and the extended
report-review-history and work-review suites. R04–R08 verified on build
`Swift_NAVFIT_0.51.0_20261101-201607_44f81547`: BuildAndTest PASS, plus the
directory-management, external-import, workspace-read-performance,
summary-groups, report-field-persistence, member-profiles and narratives suites.

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
[consolidation plan](CONSOLIDATION_PLAN.md). Existing release notes are
historical evidence; this document owns the concise current architecture.
