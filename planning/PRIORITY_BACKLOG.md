# Swift NAVFIT priority backlog

Updated 22 October 2026 from the user's review. C01/C02 implementation: r127; C03 implementation: r128 / 0.27.1-dev / schema 27. C04 is implemented in development source: fresh build and focused readiness checks pass, with a separately reproduced preexisting validation-window layout failure documented in current state; no new release is published. C05–C07 are implemented in development as 0.28.0-dev / schema 28; C08–C10 extend development to 0.29.0-dev / schema 29. Release acceptance evidence is recorded separately. It supersedes the short remaining-priorities list in the roadmap; earlier audit/batch documents remain historical evidence.

The severity below follows the user's requested categories. Reported inventory numbers (six SSN locations, approximately 100 fields, 114 globals, 267 domain-query calls, 703 raises/418 codes, 43 modules using Resume Next, and 229 form-module lines) are review observations, not a newly reproduced inventory. Verify current counts and call paths when starting each item.

## Critical — first implementation queue

| ID | Work | Scope and completion criteria |
| --- | --- | --- |
| C01 | SSN protection decision accepted | **21 October user decision:** “No, protection is covered by file folder permissions.” Retain Windows file/folder protections; do not add a database password, encryption/key setup or extra application protection requirement. SSNs remain readable in current values, revisions, recovery and retained provenance; masking is presentation only. Inventory and scope are recorded in [C01/C02 implementation notes](DIRECTORY_MERGE_R127_2026-10-21.md). No historical values or original files are scrubbed. This closes the protection-design decision; it does not claim encrypted storage. |
| C02 | Reviewed duplicate Sailor merging implemented | Schema 27 adds **Resources > Command member directory > Review possible duplicate Sailors**. Normalized surnames suggest candidates; changed-name pairs may be selected deliberately. Preview shows masked identities and report/prior IDs. Conflicting profile/report SSNs or report DoD IDs block merging. Apply requires a current preview and review reason, moves stable member links transactionally, retains the original archived profile and all report/history/trait values, and records durable mappings. Guarded Undo restores original links and metadata while retaining the audit; later affected changes block Undo. Native evidence and release scope: [C01/C02 implementation notes](DIRECTORY_MERGE_R127_2026-10-21.md). |
| C03 | Startup contributor selection implemented | **r128:** every initialized launch requires an existing contributor choice or explicit add-and-continue; nothing is preselected. First-time setup confirms its entered creator. Imported lists require a fresh choice; cancellation/reopening, stale session rejection, new-profile rollback, saves, comments and review attribution pass native checks. The context displays the selected name. Retained profiles/history remain intact. This is attribution, not authentication. See [implementation and verification scope](CONTRIBUTOR_STARTUP_R128_2026-10-21.md); old-output retesting was explicitly waived by the user for this delivery. |
| C04 | Historical and effective-dated validation implemented in development | Source-backed CH-1 calendars begin 16 December 2025, with table 1A and O5/O6 transitions from 2027; leap years, paygrades and special cases are covered. General CH-2 coverage begins 26 May 2026. **User decision: NAVADMIN 132/26 is authoritative; highest-paygrade Regular periodic exemptions begin 3 June 2026.** Earlier/missing-date coverage is explicitly Needs review, not a fabricated historical verdict. Live/saved/group/export paths share applicability; a new versioned pack preserves historical packs/results without extending legacy dates or relabeling old-engine runs. Eligibility requiring unavailable assignment/permanent-grade facts remains human review. See [current implementation and verification scope](../PROJECT_STATE.md#c04-validation-applicability). |
| C05 | Relational summary-group completion and RSCA accounting implemented in development | **Schema 28:** lifecycle, immutable posting parents, per-counted-report revision/totals, active posting pointers and unique report claims replace new XML/list accounting writes. Baseline ancestry, exact totals, NOB exclusions, corrections and retained reopen/archive accounting are relational and queryable. Production import backfills old XML once without adding a baseline; original XML/history remain evidence. Legacy posting revisions and per-report totals that were never recorded stay explicitly unknown. Native checks cover reconciliation, idempotency, cancellation preservation, duplicate/deletion/senior guards and late rollback. See [persistence contracts and verification](../PROJECT_STATE.md#c05c07-persistence-contracts). |
| C06 | Canonical field definitions and authoritative revisions implemented in development | The manifest fingerprints one versioned `field-definitions.json` catalogue for field types, controls/recovery, labels, applicability, limits/options, ordered NAVFIT/PDF bindings and fixed geometry. Services consume catalogue limits; compatibility-view generators no longer write competing VBA. `ReportHead` selects the authoritative saved revision through an ownership foreign key, replacing the circular relationship; retained report columns/pointer are checked projections. Original IDs, snapshots, changes, recovery and pending buffers remain intact. Metadata-only synthetic text/date additions exercise save/restore/checkpoint and fresh-session recovery; native fixed-page and interchange checks remain separate gates. See [implementation and verification](../PROJECT_STATE.md#c05c07-persistence-contracts). |
| C07 | Typed dates, choices and default relationships implemented in development | Five `RevisionDate` projections separate DATETIME values from Missing/Blank/NOT REQ/NOT PERF/Unknown while retaining original strings and counseling explanations. Constrained `RevisionChoice` rows hold single responses and explicit Conflict/Unknown review states; native flag pairs remain lossless compatibility projections and deliberate box selection clears the opposite box. New form/workflow domains are constrained; imported legacy exceptions stay exact. `FieldDefaultLink` verifies one target with foreign keys, retains unresolved legacy links, and blocks their application. Native persistence/transfer checks cover nulls, dates, statuses, flags, ownership and rollback; NAVFIT/PDF checks cover interchange. See [implementation and verification scope](../PROJECT_STATE.md#c05c07-persistence-contracts). |
| C08 | Implemented in schema-29 development: audit mechanisms | Map the distinct purposes of revisions/changes, report events, operation journal, folder moves, comment events and review-task events. Define consistent event identity, actor/time, operation correlation and affected-object links; remove redundant new writes only where semantics genuinely overlap. Retain domain-specific histories where needed, and preserve all older audit records, undo and review evidence. Completion requires an audited producer/consumer map and demonstrated history/rollback equivalence. |
| C09 | Implemented in schema-29 development: note/message consolidation | Confirm all current comments, replies and note drafts use `ReportMessage` and the unified report conversation UI independently of approval. Stop redundant legacy writes and retire obsolete active callers after an audit. Preserve/import `ReportNote`, `ReviewComment`, legacy mappings, authors, timestamps and thread relationships until a proven migration/retention strategy permits retirement. Test all retained schemas, idempotency, out-of-order parents, private drafts, replies and independent approval behavior. Do not reintroduce resolve/reopen requirements. |
| C10 | Implemented in schema-29 development: UI-generation and navigation consolidation | Inventory actual live destinations, generated forms, macros, control expressions, dispatchers and compatibility hosts. Retire demonstrably duplicate screens/routes and generation code, while preserving pending-edit guards and required hidden context hosts. `frmHome`, `frmLibrary` and a `frmUX*` name do not by themselves establish duplicate behavior. Prove every supported destination and return path, including reports, summary groups, My Work, notes, review, validation popouts and contextual help. Preserve the navigation rail and fixed Front/Back geometry. |

C08–C10 implementation and the producer/consumer/navigation audit are documented in [current project state](../PROJECT_STATE.md) and the [consolidation audit](CONSOLIDATION_PLAN.md#c08c10-producer-consumer-and-navigation-audit). Development verification is separate from release or physical acceptance. C01 retains the accepted file/folder-permissions decision. C04's preexisting validation-window layout gap remains separate. The user intends to simplify context/workflow requirements after C03.

## High

| ID | Work | Scope and completion criteria |
| --- | --- | --- |
| H01 | Official-template PDF report output and summary letter | Build local print-to-PDF/export using the supported fillable EVAL, CHIEFEVAL and FITREP templates as the foundation, and obtain/version/map the NAVPERS 1610/1 Summary Letter template for group summaries. Use summary-group title and Sailor name in unique, sanitized filenames; prevent collisions and overwrites and omit SSNs from filenames. Produce visibly filled, printable appearances, not merely changed XFA datasets. Verify saved-field readback, every mapped block, all seven traits, narrative capacity, continuation/overflow handling, summary roster/totals, paper geometry and actual printed fit. Keep NAVFIT98A export. Do not imply generated output constitutes official signature/submission. Decide required versus deliberately excluded fields explicitly. |
| H02 | Nonblocking NAVFIT98A export — implemented in development | 0.34.0-dev/schema 31 exports saved reports and all saved group members despite readiness findings or status, with informational findings and retained integrity safeguards. See the H02 scope below. |

H01 has an initial implementation in development source (0.30.0-dev/schema 29), with static
template output, native preview, combined/individual/Both exports and nonblocking
draft capture. [Current project state](../PROJECT_STATE.md) records the mapping,
intentional blank fields and runtime contract. Physical printed-fit acceptance and
verification on a workstation with Acrobat absent remain open; H01 is not marked
fully accepted or released. The 23 October comparison below identifies required
field-coverage and rendering corrections. Local PDF extraction in r126 remains an input feature.

**H01 correction scope, 23 October — Chief performance comments and print fidelity:**
The supplied `TEST3__Packet.pdf`, `NAVFIT98A Chief Evaluation Report.pdf` and
`NAVFIT98A Summary Letter.pdf` demonstrate differences in typography, form
geometry, placement and text capacity. The pre-correction Chief exporter incorrectly
places catalogue rating descriptions in member-specific Performance Comments
areas, omits the full performance-grade presentation, and creates additional
pages for content that fits in the NAVFIT98A example. Existing generation and
text-preservation checks did not establish semantic mapping or visual acceptance.
Preserve the supplied originals as comparison evidence; do not copy personnel
values into tracked fixtures or alter user workspaces.

**Correction implementation in development:** 0.36.0-dev/schema 33 adds the 21
authored comment fields, complete persistence/native/structured-PDF mappings,
legacy-provenance/baseline preservation and corrected grade/comment rendering.
Font-aware measurement replaces the uniform Courier estimate. See
[current implementation and verification](../PROJECT_STATE.md#h01-chief-comments-and-print-fidelity-corrections).
H01 is **not fully accepted**: physical fit and an Acrobat-absent workstation remain
open, as does acceptance of the native six-point fallback for the official
trait-34 grade-4/5 6.5-point wording. No release is published.

- **Separate marks, descriptions and authored comments.** Add 21 independent Chief
  performance-comment text values: three entries for each trait in blocks 33–39,
  matching the editable areas in the official template. Present seven trait groups
  with three comment entries each. Keep all seven traits together on the editor's
  Front Page and preserve fixed geometry, pending buffers and navigation behavior.
  Official grade descriptions remain catalogue-derived; authored comments start
  blank and must not be generated from, or overwritten by, a selected rating.
  Earlier reports receive blank new values without fabricated historical comments.
- **Complete persistence and interchange.** Incorporate the fields into the
  canonical field catalogue, saves/revisions, recovery, comparison/restoration,
  additive versioned migration and workspace transfer. Verify NAVFIT98A's actual
  comment storage before choosing its import/export mapping; preserve grouping,
  order and full text without assuming its storage matches the PDF's 21 areas.
  Cover structured-PDF input, printed output and applicable return/reconciliation
  paths. Preserve original source values, provenance and historical records.
- **Prove one corrected Chief report before broad rendering changes.** Reconcile
  the official template and NAVFIT98A geometry; correctly render grade headings,
  explanatory wording, boxed numeric marks and authored comments in their own
  areas. Respect verified field-specific fonts, sizes, alignment, line spacing and
  capacity rather than forcing every mapped value to Courier New. Inspect the
  summary sheet as well as both report pages; do not treat background differences
  as text offsets alone. Then apply verified corrections to other report kinds.
- **Evaluate the rendering route with a bounded proof.** Begin with the existing
  native Access overlay route and establish whether it can reproduce the required
  appearance reliably. Consider filling the rich PDF fields if the proof exposes
  material limitations, but require verified visible appearances, rendering and
  combined-packet assembly without Acrobat. Writing field values alone is not
  sufficient, and this backlog entry does not approve new runtime dependencies.
- **Acceptance:** compare blank, short, full and overflowing synthetic examples
  against the supplied examples and authoritative templates at actual page size.
  Verify every trait's three comments independently, rating changes, save/reopen,
  recovery, migration and applicable interchange; prove catalogue wording never
  substitutes for authored comments. Confirm correct font/placement, grade
  presentation and no avoidable continuation caused by incorrect metrics or
  geometry. Preserve genuine overflow completely and retain physical printed-fit
  and Acrobat-absent workstation acceptance. Existing nonblocking draft output,
  preview, alphabetical ordering and combined/individual/Both requirements below
  remain unchanged.

**H01 output decision, 22 October:** the user selected print-ready PDF as the
primary output, including partial drafts and incomplete reports with issues.
An editable official form is not the first delivery target.
Preserve saved revisions and official template geometry; generated output does
not create signatures or change local approval/submission state. Establish a
verified filling/rendering route before implementing the export UI around it.

**Nonblocking preview/export:** show validation findings, missing values, report
status and overflow information in the preview, but none may disable generation,
exclude a selected report or require corrective edits, approval, completion or a
status transition. No additional warning acknowledgement is required solely
because findings exist. Identify the report values being exported and preserve
pending editor/recovery values; do not force a save or discard to clear findings.
Keep missing values blank and preserve overflow text on clearly identified
continuation/additional pages without silently truncating it or changing the
official form geometry. Actual generation failures, such as an unavailable
renderer or unwritable destination, must be reported separately from report
readiness findings and must never produce a false success result.

**Combined group packet:** support one print-ready PDF containing the summary
letter/sheet first, followed by the selected reports sorted alphabetically by
Sailor last name. Use the same order for summary-sheet entries, export-preview
rows and report pages. Do not ask users to reorder reports or provide manual
reordering controls; do not infer packet order from database IDs. Resolve equal
last names consistently using first name, middle name/initial and a stable
report identity as the final tie-breaker. Keep each report's Front/Back pages and any
continuation/additional pages together. Offer the combined packet and individual
PDFs as selectable outputs, including generating both in one export run. The
application must assemble the packet itself without requiring Acrobat Pro or a
separate user-operated PDF merge tool. The chosen generation/assembly route must
be verified on the supported workstation environment without Acrobat installed;
developer-only tools must not become undeclared end-user dependencies.

### H02 — Nonblocking NAVFIT98A export

Implemented in **0.34.0-dev/schema 31** development source. See
[behavior and verification scope](../PROJECT_STATE.md#h02-nonblocking-navfit98a-export).
No release or publication is included. H01's print-fidelity corrections remain open.

**23 October user decision:** allow NAVFIT98A export of reports and summary groups
regardless of existing validation findings, missing/incomplete values, readiness
or report/group status. Inform users that validation issues exist, with access to
the findings, but do not require corrections, status changes or an additional
acknowledgement solely because findings exist. Do not omit reports from a group
export because of validation findings. Export does not imply approval, signature
or readiness for submission.

Audit report/group export entry points and shared export services so the same
nonblocking policy applies consistently. Preserve pending editor/grid values and
existing explicit save/capture behavior; do not silently save or discard drafts
to bypass validation. Keep actual export-integrity and operational safeguards:
source/revision consistency, valid serialization, destination access, collision
protection and truthful failure reporting. Distinguish these failures from report
readiness findings; never produce a corrupt file or claim success after failure.

Verify report and summary-group exports with clean and issue-bearing synthetic
data, including incomplete drafts and mixed groups. Confirm all selected reports
are retained, informational findings do not prevent export, source/history remain
unchanged apart from deliberate user saves, and actual generation failures still
report failure. Retain NAVFIT98A field coverage, provenance and round-trip checks.

## Medium

| ID | Work | Scope and completion criteria |
| --- | --- | --- |
| M01 | Compiled production distribution, ideally signed | Produce a production `.accde` and verify it on the supported desktop Access architecture/version. Continue publishing Blank and synthetic Demo `.accdb` files for testing. Establish signing/certificate and release-integrity procedures when an approved certificate is available; verify the actual distributed files, startup, imports and provenance. ACCDE/source signing is hardening, not SSN encryption or an authorization boundary. Free Access Runtime support remains out of scope. Reliable fresh generation is a prerequisite (RG01). |
| M02 | Persist or explicitly own wizard state | Inventory module-level state, including `modNAVFITExchange` returned path/values/revision and other import/return previews. Use an explicit context object for bounded in-session ownership and staging tables where recovery across errors/resets is required. Record source fingerprints, target revision and choices; support deliberate resume/cancel/cleanup and reject stale previews. Test VBA reset, unhandled failure, restart and source/target changes without silent data loss or accidental Apply. Protect staged identifiers under C01. |
| M03 | Replace query-per-row patterns with joined/batched reads | Inventory and profile domain queries, especially loops in `modBulkFolders` and report/group/directory displays. Reuse joined/batched presentation data and refresh-scoped facts while retaining authoritative write-time revision/permission/state checks. Benchmark representative large synthetic workspaces on exact artifacts; prove equivalent Null, NOB, identity, RSCA, recovery and history behavior. Prior bulk import/verification improvements remain delivered and are not a reason to relax preservation checks. |
| M04 | Preserve useful error diagnostics and narrow suppression | Inventory reused error codes and all `On Error Resume Next` sites. Give failures a stable operation/error identity plus module/procedure, stage, original number/description and transaction outcome; prevent cleanup from masking the primary error. Keep suppressed cleanup narrowly scoped and restore handling immediately. Separate event-expression adapters from domain methods where useful; Boolean callbacks are not automatically defects. Test rollback, retry, UI feedback and redacted diagnostics without logging SSNs. |
| M05 | Reduce brittle form/control string coupling | Define explicit form/view contracts, centralized identifiers or typed adapters, and validate required controls/events at build/startup boundaries. Fail clearly on missing/renamed controls rather than silently skipping behavior. Audit generators, macros and expression callbacks as well as VBA. Prove rename/missing-control detection and navigation/save/recovery behavior; form-module line count alone is not an acceptance measure. |
| M06 | Store the trait catalogue as versioned runtime data | Move runtime lookup to `TraitDefinition` (form kind, block, mark, version/source), with uniqueness/relationships and exact official wording. Retain the existing JSON catalogue as a reviewed seed/generation source or consolidate it into the agreed canonical field/catalogue process. Remove redundant generated `Select Case` text and concatenation noise after equivalence is proven. Normalize display title styling deliberately without changing authoritative text. Compare every trait/mark/source reference before retiring the current generator. |
| M07 | Establish accessibility acceptance | Create and execute keyboard, focus/tab-order, screen-reader labels/announcements, contrast/high-contrast, display/DPI and scaling acceptance for the actual supported workflows/artifacts. Research applicable Section 508 requirements and document traceable results, defects and remediation; the present absence of acceptance is a risk, not a completed compliance determination. Include startup contributor selection, dialogs/popouts, report tools, grids, findings and output. Automated geometry/binding checks do not establish accessibility. Physical review remains deferred until the user can participate. |

M01 has development tooling and native compiled-candidate proof on desktop Access
16.0.20430.20182 x64, via the existing development entry point. It is **not complete
for production**: normal-policy launch remains unresolved after cancelling an
unsigned-file security prompt, no approved
signing certificate is available, and compiled PDF output remains unverified.
The ACCDE import-picker filter is implemented in 0.37. See
[current scope and limitations](../PROJECT_STATE.md#environment-build-verification-and-delivery)
and the [compiled-candidate procedure](../docs/Development_Workflow.md#compiled-desktop-candidate-m01).
No release/publication or security-policy change is included.

M02 is implemented in **0.31.0-dev/schema 30** development source: contributor-owned
durable descriptors, deliberate Resume/Cancel, source/target guards, atomic stage
consumption and preserved preview rows during workspace import. The inventory and
ephemeral-versus-durable policy are in [current project state](../PROJECT_STATE.md#m02-preview-ownership-and-recovery).
Development/reset/migration checks remain separate from release acceptance.

M03 is implemented in **0.32.0-dev/schema 30** development source: chunked folder
reads and Undo prechecks, refresh-owned paths, batched directory-history averages
and shared scalar/batch group persistence validation. Existing prepared Reports/
My Work readers and live write guards remain. The [inventory and verification
scope](../PROJECT_STATE.md#m03-joinedbatched-read-paths) records measured native
equivalence/performance. The preexisting read-row fixture failure is repaired in
0.38 without weakening cache-divergence rejection. No schema or release change is included.

M04–M06 implementation is in **0.33.0-dev/schema 31** development source:
contextual primary/cleanup diagnostics and narrower suppression; required
form/control/event contracts and installed form/report/macro verification; and
versioned runtime trait reference data seeded from the reviewed catalogue.
See [current contracts and verification](../PROJECT_STATE.md#m04m06-diagnostics-view-contracts-and-runtime-trait-data)
and the [source/consumer audit](CONSOLIDATION_PLAN.md#m04m06-source-and-consumer-audit).
M04–M06 did not implement H01's Chief performance-comment/rendering corrections or
H02. H02 is implemented separately above; no release is published.

## Low

| ID | Work | Scope and completion criteria |
| --- | --- | --- |
| L01 | Descriptive foreign-key names — implemented in development | Schema 32 uses descriptive names for all 34 formerly opaque major relationships. Production Settings import retains original fields/IDs and unchanged relationship enforcement/attributes in the fresh target; source schemas remain untouched. See the implementation scope below. |
| L02 | Clarify revision-ID field names — implemented in development | Writable schema-owned query aliases expose `ExpectedRevisionID` and `BaseRevisionID`; active batch/meeting/comparison consumers use them. Legacy physical columns remain the sole storage for exact historical import, without duplicated values or renumbering. See the implementation scope below. |

L01–L02 are implemented in **0.35.0-dev/schema 32** development source. See
[behavior, compatibility and verification](../PROJECT_STATE.md#l01l02-relationship-names-and-revision-id-aliases).
No release or publication is included; historical fixture waivers remain in force.

**Remote follow-up, 0.37.0-dev/schema 33:** repaired the stale Members/workbench
fixtures without dropping their behavioral assertions, fixed non-grid proposal
projection and saved-export entry guards found by reordered task tests, and
added native label/focus/resize checks. M07 has development preparation and
targeted label/layout remediation, not accessibility acceptance. See
[current scope and verification](../PROJECT_STATE.md#remote-regression-and-accessibility-preparation).

**Reliability follow-up, 0.38.0-dev/schema 33:** current-schema ACCDB/ACCDE
backup/restore, protected verification files and stale-token cancellation for PDF/
native-return previews. Focused fixtures now honor contributor startup and reject
deliberately divergent report caches before unrelated mutations. Native test and
retained-schema evidence is recorded in the same current-state section; historical
fixture waivers and attended acceptance remain unchanged.
**M02/M03/M04 follow-ups, 0.39.0-dev/schema 33:** workspace-import cancellation
is token-guarded and reopened previews require deliberate Resume. All three
cancel adapters retain stale choices with recovery guidance. My Work's redundant
closed-window return refresh (reproduced on 0.37/0.38) is removed while preserving
initialization, activation and loaded-window refreshes. Native report/group export
and retained-source rebuild preserve primary errors across cleanup/journal failures,
redact diagnostic identifiers and disclose file-publication state. Existing focused
runners cover failure/retry and post-publication journal faults; no schema change,
new dependency, release or attended acceptance is implied.

## Initial source checks and partial delivery

- `modCommandDirectory.DirectoryIdentityMatch` explicitly treats SSN as authoritative, then uses a unique normalized full-name fallback that excludes conflicting SSNs. C02 now adds deliberate surname candidates and reviewed merging; automatic name matching was not broadened.
- C04 replaces the pre-2026 calendar cutoff and O5/O6 year branches with verified calendar versions; CMDCS cannot bypass coverage selection. Missing/unsupported general coverage now emits an explicit Needs review finding even on otherwise clean reports.
- The legacy draft pack in `modValidation` still ends 3 October 2026. New saved runs select the applicability-v1 shared-readiness pack in `modSavedReadiness`; its open end does not assert complete historical policy coverage. Historical r110 shared results remain readable, but rerunning that engine under the updated implementation is explicitly rejected. Historical dates/results were not rewritten.
- `modReportMessages` already defines one report-level conversation system with preserved legacy mappings. C09 is completion/retirement work, not a request to create another messaging model.
- The current trait procedure is generated from `src/trait-catalogue.json` by `scripts/write-trait-catalogue.ps1`. M06 addresses its runtime representation, duplication and presentation; it is not an assertion that the catalogue has no existing data source.
- Retired standalone help/review routes and source/form-binding verification have already been implemented in their documented scope. C10/M05 require a current caller audit before removing further code or forms.

## Existing work retained and release gates

- **RG01 — Current fresh-generation proof obtained, 21 October:** the complete source-only r127 build passes all 71 stages, compile/reopen, 99 static modules and bindings across 66 forms, with native workflow and migration verification against the fresh blank. Use the standard source-first workflow; see [setup evidence](DEVELOPMENT_SETUP_2026-10-21.md). The earlier r126 closed-object failure did not reproduce and has no established root cause. Retain its historical evidence and capture diagnostics if it returns; keep it distinct from the earlier transient migration-session crash.
- Preserve the existing NAVFIT98A field-coverage, staged return/reconciliation and round-trip proof work. H01 complements it; it does not remove export compatibility.
- Continue source-backed grouping exceptions, completeness and ready-for-signature guidance under C04, plus focused usability/performance work under C10/M03/M05.
- Continuity history, counseling support and local submission tracking remain later proposals pending agreed scope. Local review/group completion remains separate from official submission and signatures.
- Every implementation release must preserve source workspaces and earlier artifacts; use additive, versioned migrations; retain original IDs, relationships, revisions, reviews, notes, provenance and recovery; and pass retained-schema, populated-transfer, source-hash, rollback and rejection gates. Privacy/normalization changes do not authorize destructive history cleanup.
- Maintain fixed report pages, all seven Front Page traits, the left navigation rail and pending-buffer/recovery behavior. Publish newly named, verified artifacts under `output`, verify required Drive uploads, and close owned Access sessions after work. A planning-only update requires no new application build or upload.

See [current state](CURRENT_STATE.md) for delivered evidence and [roadmap](ROADMAP.md) for the overview.
