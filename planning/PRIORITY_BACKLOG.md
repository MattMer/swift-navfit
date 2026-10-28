# Swift NAVFIT priority backlog

Updated 22 October 2026 from the user's review. C01/C02 implementation: r127; C03 implementation: r128 / 0.27.1-dev / schema 27. C04 is implemented in development source: fresh build and focused readiness checks pass, with a separately reproduced preexisting validation-window layout failure documented in current state; no new release is published. C05–C07 are implemented in development as 0.28.0-dev / schema 28; C08–C10 extend development to 0.29.0-dev / schema 29. Release acceptance evidence is recorded separately. It supersedes the short remaining-priorities list in the roadmap; earlier audit/batch documents remain historical evidence.

The severity below follows the user's requested categories. Reported inventory numbers (six SSN locations, approximately 100 fields, 114 globals, 267 domain-query calls, 703 raises/418 codes, 43 modules using Resume Next, and 229 form-module lines) are review observations, not a newly reproduced inventory. Verify current counts and call paths when starting each item.

## Critical — first implementation queue

| ID | Work | Scope and completion criteria |
| --- | --- | --- |
| C01 | SSN protection decision accepted | **21 October user decision:** “No, protection is covered by file folder permissions.” Retain Windows file/folder protections; do not add a database password, encryption/key setup or extra application protection requirement. SSNs remain readable in current values, revisions, recovery and retained provenance; masking is presentation only. Inventory and scope are recorded in [C01/C02 implementation notes](DIRECTORY_MERGE_R127_2026-10-21.md). No historical values or original files are scrubbed. This closes the protection-design decision; it does not claim encrypted storage. |
| C02 | Reviewed duplicate Sailor merging implemented | Schema 27 adds **Resources > Command member directory > Review possible duplicate Sailors**. Normalized surnames suggest candidates; changed-name pairs may be selected deliberately. Preview shows masked identities and report/prior IDs. Conflicting profile/report SSNs or report DoD IDs block merging. Apply requires a current preview and review reason, moves stable member links transactionally, retains the original archived profile and all report/history/trait values, and records durable mappings. Guarded Undo restores original links and metadata while retaining the audit; later affected changes block Undo. Native evidence and release scope: [C01/C02 implementation notes](DIRECTORY_MERGE_R127_2026-10-21.md). |
| C03 | Startup contributor selection implemented | **r128:** every initialized launch requires an existing contributor choice or explicit add-and-continue; nothing is preselected. First-time setup confirms its entered creator. Imported lists require a fresh choice; cancellation/reopening, stale session rejection, new-profile rollback, saves, comments and review attribution pass native checks. The context displays the selected name. Retained profiles/history remain intact. This is attribution, not authentication. See [implementation and verification scope](CONTRIBUTOR_STARTUP_R128_2026-10-21.md); old-output retesting was explicitly waived by the user for this delivery. |
| C04 | Historical and effective-dated validation implemented in development | Source-backed CH-1 calendars begin 16 December 2025, with table 1A and O5/O6 transitions from 2027; leap years, paygrades and special cases are covered. General CH-2 coverage begins 26 May 2026. **User decision: NAVADMIN 132/26 is authoritative; highest-paygrade Regular periodic exemptions begin 3 June 2026.** Earlier/missing-date coverage is explicitly Needs review, not a fabricated historical verdict. Live/saved/group/export paths share applicability; a new versioned pack preserves historical packs/results without extending legacy dates or relabeling old-engine runs. Eligibility requiring unavailable assignment/permanent-grade facts remains human review. See [current implementation and verification scope](PROJECT_STATE_HISTORY.md#c04-validation-applicability). |
| C05 | Relational summary-group completion and RSCA accounting implemented in development | **Schema 28:** lifecycle, immutable posting parents, per-counted-report revision/totals, active posting pointers and unique report claims replace new XML/list accounting writes. Baseline ancestry, exact totals, NOB exclusions, corrections and retained reopen/archive accounting are relational and queryable. Production import backfills old XML once without adding a baseline; original XML/history remain evidence. Legacy posting revisions and per-report totals that were never recorded stay explicitly unknown. Native checks cover reconciliation, idempotency, cancellation preservation, duplicate/deletion/senior guards and late rollback. See [persistence contracts and verification](PROJECT_STATE_HISTORY.md#c05c07-persistence-contracts). |
| C06 | Canonical field definitions and authoritative revisions implemented in development | The manifest fingerprints one versioned `field-definitions.json` catalogue for field types, controls/recovery, labels, applicability, limits/options, ordered NAVFIT/PDF bindings and fixed geometry. Services consume catalogue limits; compatibility-view generators no longer write competing VBA. `ReportHead` selects the authoritative saved revision through an ownership foreign key, replacing the circular relationship; retained report columns/pointer are checked projections. Original IDs, snapshots, changes, recovery and pending buffers remain intact. Metadata-only synthetic text/date additions exercise save/restore/checkpoint and fresh-session recovery; native fixed-page and interchange checks remain separate gates. See [implementation and verification](PROJECT_STATE_HISTORY.md#c05c07-persistence-contracts). |
| C07 | Typed dates, choices and default relationships implemented in development | Five `RevisionDate` projections separate DATETIME values from Missing/Blank/NOT REQ/NOT PERF/Unknown while retaining original strings and counseling explanations. Constrained `RevisionChoice` rows hold single responses and explicit Conflict/Unknown review states; native flag pairs remain lossless compatibility projections and deliberate box selection clears the opposite box. New form/workflow domains are constrained; imported legacy exceptions stay exact. `FieldDefaultLink` verifies one target with foreign keys, retains unresolved legacy links, and blocks their application. Native persistence/transfer checks cover nulls, dates, statuses, flags, ownership and rollback; NAVFIT/PDF checks cover interchange. See [implementation and verification scope](PROJECT_STATE_HISTORY.md#c05c07-persistence-contracts). |
| C08 | Implemented in schema-29 development: audit mechanisms | Map the distinct purposes of revisions/changes, report events, operation journal, folder moves, comment events and review-task events. Define consistent event identity, actor/time, operation correlation and affected-object links; remove redundant new writes only where semantics genuinely overlap. Retain domain-specific histories where needed, and preserve all older audit records, undo and review evidence. Completion requires an audited producer/consumer map and demonstrated history/rollback equivalence. |
| C09 | Implemented in schema-29 development: note/message consolidation | Confirm all current comments, replies and note drafts use `ReportMessage` and the unified report conversation UI independently of approval. Stop redundant legacy writes and retire obsolete active callers after an audit. Preserve/import `ReportNote`, `ReviewComment`, legacy mappings, authors, timestamps and thread relationships until a proven migration/retention strategy permits retirement. Test all retained schemas, idempotency, out-of-order parents, private drafts, replies and independent approval behavior. Do not reintroduce resolve/reopen requirements. |
| C10 | Implemented in schema-29 development: UI-generation and navigation consolidation | Inventory actual live destinations, generated forms, macros, control expressions, dispatchers and compatibility hosts. Retire demonstrably duplicate screens/routes and generation code, while preserving pending-edit guards and required hidden context hosts. `frmHome`, `frmLibrary` and a `frmUX*` name do not by themselves establish duplicate behavior. Prove every supported destination and return path, including reports, summary groups, My Work, notes, review, validation popouts and contextual help. Preserve the navigation rail and fixed Front/Back geometry. |

C08–C10 implementation and the producer/consumer/navigation audit are documented in [current project state](../PROJECT_STATE.md) and the [consolidation audit](CONSOLIDATION_PLAN.md#c08c10-producer-consumer-and-navigation-audit). Development verification is separate from release or physical acceptance. C01 retains the accepted file/folder-permissions decision. C04's earlier validation-window layout gap was addressed in 0.37 (see PROJECT_STATE.md). The user intends to simplify context/workflow requirements after C03.

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
The 0.42 UI correction sizes each Chief authored comment like the printed
comment box and places read-only selected-grade wording under its trait title.
Summary-list actions open PDF preview/output for the selected group. These
editing/navigation changes do not establish printed-fit acceptance.
Font-aware measurement replaces the uniform Courier estimate. See
[current implementation and verification](PROJECT_STATE_HISTORY.md#h01-chief-comments-and-print-fidelity-corrections).
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
[behavior and verification scope](PROJECT_STATE_HISTORY.md#h02-nonblocking-navfit98a-export).
No release or publication is included. H01's print-fidelity corrections remain open.

The 0.42 user flow exports the selected entire summary group from its Members
page or the summary list, with an internal verified native template and one
export-directory choice. Returned-file/template/report selectors are not shown.
Existing native provenance, baseline and reconciliation services remain retained.

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
| M03 | Replace query-per-row patterns with joined/batched reads | Inventory and profile domain queries, especially loops in `modBulkFolders` and report/group/directory displays. Reuse joined/batched presentation data and refresh-scoped facts while retaining authoritative write-time revision/permission/state checks. Benchmark representative large synthetic workspaces on exact artifacts; prove equivalent Null, NOB, identity, RSCA, recovery and history behavior. Prior bulk import/verification improvements remain delivered and are not a reason to relax preservation checks.<br>**Progress 27 Oct 2026** (synthetic 4,700-report workspace): saves 1.9 s → 43 ms and new reports 3.6 s → 70 ms (`CommitDraft` append-only writes); Reports list 2.7–4.5 s → 0.41 s and a 250-member group grid 3.2–4.8 s → 1.15 s (indexed seeks of saved traits, one PSG pass, joins instead of nested `IN`). **Still open, revisit later (user decision):** editor open about 1 s, flat with workspace size and mostly page layout; skipping unchanged property writes made it slower, and saving the field-page controls hidden by default broke page rendering. A 250-report group PDF packet takes about 110 s, 96 s of it in `OutputTo` (the form image is about 25%). |
| M04 | Preserve useful error diagnostics and narrow suppression | Inventory reused error codes and all `On Error Resume Next` sites. Give failures a stable operation/error identity plus module/procedure, stage, original number/description and transaction outcome; prevent cleanup from masking the primary error. Keep suppressed cleanup narrowly scoped and restore handling immediately. Separate event-expression adapters from domain methods where useful; Boolean callbacks are not automatically defects. Test rollback, retry, UI feedback and redacted diagnostics without logging SSNs.<br>**Found 27 Oct 2026:** `UXFixedSelect` (`modReportPages`) turns on `On Error Resume Next` before reading `Err`, so every page failure reads "Unable to open report page: 0" with no description. |
| M05 | Reduce brittle form/control string coupling | Define explicit form/view contracts, centralized identifiers or typed adapters, and validate required controls/events at build/startup boundaries. Fail clearly on missing/renamed controls rather than silently skipping behavior. Audit generators, macros and expression callbacks as well as VBA. Prove rename/missing-control detection and navigation/save/recovery behavior; form-module line count alone is not an acceptance measure. |
| M06 | Store the trait catalogue as versioned runtime data | Move runtime lookup to `TraitDefinition` (form kind, block, mark, version/source), with uniqueness/relationships and exact official wording. Retain the existing JSON catalogue as a reviewed seed/generation source or consolidate it into the agreed canonical field/catalogue process. Remove redundant generated `Select Case` text and concatenation noise after equivalence is proven. Normalize display title styling deliberately without changing authoritative text. Compare every trait/mark/source reference before retiring the current generator. |
| M08 | Consolidate test runners into one suite per feature | Using [tests/runners.json](../tests/runners.json), merge runners that test the same feature from different milestones into one suite per area, keeping every check that still holds. Resolve each `pending` entry: fix it, recreate its fixture from tracked definitions, or delete it with the feature evidence. Done when every runner is `core`, `release`, `extended`, `manual` or `helper`, the extended tier passes, and no area has more than a few runners. Start after the 24 Oct 2026 diagnosis of 18 failing runners is merged.<br>**Stale manual runners found 27 Oct 2026:** `test-batched-reads` (`ReadBenchmarkHarness` opens the retired `frmBulkFolders`); `test-persistence-transfer` requires its second build to be schema 28. |
| M07 | Establish accessibility acceptance | Create and execute keyboard, focus/tab-order, screen-reader labels/announcements, contrast/high-contrast, display/DPI and scaling acceptance for the actual supported workflows/artifacts. Research applicable Section 508 requirements and document traceable results, defects and remediation; the present absence of acceptance is a risk, not a completed compliance determination. Include startup contributor selection, dialogs/popouts, report tools, grids, findings and output. Automated geometry/binding checks do not establish accessibility. Physical review remains deferred until the user can participate. |
| M09 | Resize and repaint performance (one grouped change) — implemented in development | **User decision 26 Oct 2026:** keep resizable workspace screens (Reports, Summary groups, My Work, Resources, Settings, Recovery, Snippets) and the fixed report editor. Work these together:<br>**Speed, no visible change:**<br>(1) An unchanged size skips all work in `UXResizeWorkspace`. Today a cache hit still re-runs sort headers, help layout and the palette.<br>(2) Paint the palette once when a form opens, and repaint status labels only when their text changes. Remove the duplicate `UXApplyVisual` in `UXHelpResize` and the extra repaint from `OnActivate` and `UXLoadContext`.<br>(3) Turn form `Painting` off during layout.<br>**Fixed where sizing adds nothing:**<br>(4) *Dropped:* the pop-up dialogs stay resizable. Many are 9–12 inches wide and up to 10 inches tall, so on a laptop at 150% scaling resizing and scrolling is the only way to reach every control. Their resize handler only moves the help button.<br>(5) The RSCA table and (6) the trait grid shrink to fit narrow windows but no longer stretch past their designed widths.<br>**Done in development:** items 1–3, 5–6, and 7–8 without tags: on windows wider than the design, lists, subforms, wide text boxes and wide captions or outlines stretch by their share of the extra width. Controls to their right move over, everything else keeps its design size, and content stops at 16,500 twips. Narrower windows still scale proportionally. Workspace layout also caps widths at Access's 31,680-twip form limit; past it, the resize cleanup used to loop forever (a maximized window on a 2560-pixel or scaled 4K screen). Item 9 done: the rail, content, compact and design numbers are named constants in `modUXViewport`. Per-screen passes retired: the empty `modUXReviewLayout` was removed, and `UXWorkLayout` was replaced by shared list-column scaling (resize and `UXRestoreViewport`) plus corrected My Work design column widths. The bulk-field reflow now skips resizes that change nothing. `UXValidationLayout` (pop-up), `UXHeaderReviewLayout` (fixed editor) and `UXGridViewport` (trait grid) remain, because they lay out their own content. **M09 is complete in development.** Skipping write-backs of unchanged layout values, caching SQL list headings and suspending `Painting` during layout roughly halved new-size resize time and cut same-size resizes by up to 80%. The palette already skipped unchanged writes (3–8 ms), so painting it once per form isn't needed.<br>**Layout quality (show before-and-after screenshots for user review):**<br>(7) Only elastic controls stretch: lists, text areas, search boxes. Buttons keep their designed width, and right-hand actions stay pinned right. A per-control layout tag replaces the per-screen second passes (`UXWorkLayout`, `UXReviewLayout`, `UXHeaderReviewLayout`, `UXValidationLayout`).<br>(8) Cap the content width on very wide windows.<br>(9) Move the rail, design-width and compact-breakpoint numbers (2450, 2850, 10450, 12000) into named constants.<br>**Acceptance:** measure open, activate and maximize times before and after on the main screens. Viewport, shell, visual and accessibility runners pass. Remove the tests for any removed resize behavior in the same commit. The report editor's fixed geometry is unchanged. |
| M10 | Architecture description (DoDAF 2.02 fit-for-purpose) — implemented in development | **Steps 1–2 done in development:** [docs/architecture](../docs/architecture/Swift_NAVFIT_Architecture_Description.md) holds the 42010-structured description and 14 views. AV-2, DIV-2, DIV-3, OV-6a, SV-8 and the AV-1/SV-7 fact blocks are generated by `scripts/generate-architecture-views.ps1`. The extended runner `test-architecture-views` fails when they drift from the build or when the OV-6b states differ from the schema. **Step 3 done:** the [DevSecOps alignment matrix](../docs/architecture/DevSecOps_Alignment.md) assesses 118 guidebook activities: 52 done, 27 partial, 5 gaps, 14 command-owned, 20 not applicable. **Step 4 done:** a generated [SBOM](../docs/architecture/SBOM.md) (Markdown and CycloneDX 1.5, checked by `test-architecture-views`), a STRIDE [threat model](../docs/architecture/Threat_Model.md) with six recommendations, and the code-signing plan (threat model section 5). The matrix now shows 54 done, 28 partial, 2 gaps. **M10 complete in development.** Malformed-input importer tests are done: `test-malformed-imports`, 55 cases, no defects found. `VBScript.RegExp` is replaced in 0.49.0, verified against the old code on 20,000 random strings. Remaining follow-up: the signing rehearsal (M01).<br>**User request 26 Oct 2026:** present a report aligned with DoDAF 2.02, as we do for Section 508. References are in `FORMS_REFERENCES_TEMPLATES/Architecture/` (DoDAF 2.02, the ISO/IEC/IEEE 42010 architecture-description template, the DevSecOps Activities & Tools Guidebook tables, and the DoD CIO open-source software memo).<br>**Build in this order:**<br>(1) An architecture description in the 42010 structure: stakeholders and concerns, chosen viewpoints and rationale, and the views AV-1, OV-1, OV-4, OV-5b, OV-6b, SV-1, SV-4, SV-8 and StdV-1.<br>(2) AV-2, DIV-2/3, OV-6a and SV-8 generated from tracked sources (schema definition, field catalogue, validation rule packs, migration history), with a check that fails when they go stale.<br>(3) A DevSecOps alignment matrix in the guidebook's tailoring columns: each activity done (naming the script or receipt), not applicable (with reason), or a gap.<br>(4) Close the cheap gaps: an SBOM (VBA references, Office/Access version, test-only tools), a lightweight threat model for a local file-based app, and the M01 code-signing plan.<br>**Not applicable:** Services viewpoint (no network services); most Capability and Project models (no acquisition program); network deployment, operations and monitoring activities.<br>**Outside our control:** formal DoDAF conformance (DM2 data and PES exchange need a DoD architecture tool); repository registration, ATO/RMF, mission cyber risk assessment and interoperability certification (command and ISSM). The report states it is DoDAF-aligned, not certified conformant. |
| M11 | Optional SharePoint Lists workspace storage — **proposed, planning only** | **User request 26 Oct 2026.** Each command chooses a storage mode:<br>**Local** (the default and current behavior): reports and supporting data stay in the Swift NAVFIT `.accdb`.<br>**Shared:** each user runs a local Access front end linked to the command's SharePoint Lists.<br>The report editor, validation, summary groups, rack and stack, comments, revisions, routing, and NAVFIT98A import/export stay available in both modes.<br>Requirements, constraints, open decisions and phasing are in the [M11 section](#m11--optional-sharepoint-lists-workspace-storage) below. Nothing is implemented, and current behavior is unchanged. |

M01 has development tooling and native compiled-candidate proof on desktop Access
16.0.20430.20182 x64, via the existing development entry point. It is **not complete
for production**: normal-policy launch remains unresolved after cancelling an
unsigned-file security prompt, and no approved
signing certificate is available. Compiled PDF output (report PDF and group packet
with summary sheet) is verified through the production dialog by the compiled
runner on Access 16.0.20527.20088 x64; printed fit and an Acrobat-absent workstation remain H01 acceptance.
The ACCDE import-picker filter is implemented in 0.37. See
[current scope and limitations](PROJECT_STATE_HISTORY.md#environment-build-verification-and-delivery)
and the [compiled-candidate procedure](../docs/Development_Workflow.md#compiled-desktop-candidate-m01).
No release/publication or security-policy change is included.

M02 is implemented in **0.31.0-dev/schema 30** development source: contributor-owned
durable descriptors, deliberate Resume/Cancel, source/target guards, atomic stage
consumption and preserved preview rows during workspace import. The inventory and
ephemeral-versus-durable policy are in [current project state](PROJECT_STATE_HISTORY.md#m02-preview-ownership-and-recovery).
Development/reset/migration checks remain separate from release acceptance.

M03 is implemented in **0.32.0-dev/schema 30** development source: chunked folder
reads and Undo prechecks, refresh-owned paths, batched directory-history averages
and shared scalar/batch group persistence validation. Existing prepared Reports/
My Work readers and live write guards remain. The [inventory and verification
scope](PROJECT_STATE_HISTORY.md#m03-joinedbatched-read-paths) records measured native
equivalence/performance. The preexisting read-row fixture failure is repaired in
0.38 without weakening cache-divergence rejection. No schema or release change is included.

M04–M06 implementation is in **0.33.0-dev/schema 31** development source:
contextual primary/cleanup diagnostics and narrower suppression; required
form/control/event contracts and installed form/report/macro verification; and
versioned runtime trait reference data seeded from the reviewed catalogue.
See [current contracts and verification](PROJECT_STATE_HISTORY.md#m04m06-diagnostics-view-contracts-and-runtime-trait-data)
and the [source/consumer audit](CONSOLIDATION_PLAN.md#m04m06-source-and-consumer-audit).
M04–M06 did not implement H01's Chief performance-comment/rendering corrections or
H02. H02 is implemented separately above; no release is published.

### M11 — Optional SharePoint Lists workspace storage

**Status:** proposed and not scheduled; this entry is planning only. Today Swift
NAVFIT is single-file and offline: one authoritative workspace `.accdb`, one
contributor at a time ([PROJECT_STATE.md](../PROJECT_STATE.md)), and no network
interfaces ([SV-1](../docs/architecture/views/SV-1.md)). Those statements stay true
until this work ships. No other backlog item covers shared or multi-user storage.

**Objective:** each command chooses a storage mode.
- **Local mode (default):** reports and supporting data remain inside the Swift
  NAVFIT `.accdb`, exactly as today.
- **Shared mode:** each user runs a local Access front end connected to the
  command's SharePoint Lists.

The report editor, validation, summary groups, rack and stack, comments,
revisions, routing, and NAVFIT98A import/export remain available in both modes.

**Requirements:**

1. **Setup/settings screen.** A workspace setup and settings screen shows the
   active command, the workspace, and the storage mode.
2. **Discovery.** For SharePoint mode, the user enters a site URL and Swift
   discovers an existing Swift NAVFIT workspace on that site. Users never choose
   individual lists.
3. **Provisioning.** If the workspace doesn't exist, Swift offers guided
   provisioning of the required lists, columns, indexes, configuration records
   and supported settings, subject to the user's SharePoint permissions.
4. **Detection and recovery.** Swift detects the schema version, incompatible
   structures, missing components and insufficient permissions. Interrupted
   provisioning can be resumed, and never overwrites existing command data.

**Product invariants that still apply** ([AGENTS.md](../AGENTS.md)):
- Every release imports every earlier schema, in either mode.
- Schema changes are additive and versioned in both modes, and the SharePoint
  schema carries its own version, checked against the front end's.
- Report history is never deleted, except confirmed draft disposal.
- A populated workspace is never silently replaced, and records are never merged
  by ID. Moving between modes uses the production import into an empty target.
- The fixed report-editor geometry, the unbound buffer and recovery are unchanged.
  Pending-edit recovery checkpoints stay local to each user's front end.
- Demo and test data stay synthetic. Never put personnel data on a test site.

**Constraints of SharePoint Lists that the design must address:**

| Area | Today (local `.accdb`) | With SharePoint Lists |
| --- | --- | --- |
| Transactions | A save, its revision, changes and receipt commit together | No transaction across lists or items. Needs an idempotent write protocol (a commit marker or journal per save) and reconciliation after interruption. |
| Integrity rules | Table validation rules, required fields, unique indexes, enforced relationships | Not enforced the same way: limited unique columns, lookup relationships, no table-level rules. Swift must enforce them in its data layer and verify on read. |
| Identity | AutoNumber IDs are referenced throughout | SharePoint assigns its own item IDs. Use Swift's existing GUIDs as stable keys and map IDs. |
| Size and indexing | No practical row limit for a binder | List view threshold (5,000 items) and per-list index limits. Queries must stay on indexed columns, with archiving guidance. |
| Concurrency | One contributor at a time | Several users at once. Optimistic concurrency on every write (extending the existing version counters, or using item versions) with clear conflict messages. |
| Attribution | Contributor choice; attribution only (decisions C01, C03) | SharePoint records an authenticated "Modified By". Decide whether the contributor follows the signed-in account. |
| Performance | Local disk | Network latency. Batched reads (M03) and caching matter more. Define offline behavior. |
| Security | Folder permissions on approved storage (C01) | SharePoint permissions, plus authorization of the tenant and site for PII/CUI. Decide whether SSNs are stored there at all. |

**Decisions (product owner, 26 Oct 2026):**

| Topic | Decision |
| --- | --- |
| Environment | DoD365 SharePoint Online, the command's site |
| SSNs | Stored in SharePoint if the ISSM approves; otherwise shared mode keeps them out |
| Offline | Read-only cache: users can view cached data but not save until reconnected |
| Attribution | In shared mode, the contributor is the signed-in account; local mode keeps the startup profile choice |
| Moving between modes | Both directions (local to shared, shared to local), each through the verified import into an empty target |
| Provisioning | From inside Swift's setup screen, using the method the research spike proves reliable on DoD365 |
| Who may set up or migrate | SharePoint site owners (list-management rights); Swift checks those permissions first |

**Still open (needs the ISSM):**
- Confirmation that the command's DoD365 site is authorized for this PII/CUI.
- Whether SSNs may be stored there. Until approved, plan for the no-SSN variant.

**Design notes for planning:**
- A **storage-provider layer** behind one interface, so the editor, validation,
  groups, routing, comments and NAVFIT98A exchange don't change with the mode.
  Local mode keeps the current tables. Shared mode uses linked SharePoint lists.
- **A configuration list** marks a Swift NAVFIT workspace. It holds the workspace
  GUID, command, schema version, minimum front-end version and a provisioning
  journal. Discovery looks for it, so users never pick lists.
- **Provisioning is resumable.** Create components in dependency order and record
  each completed step in the journal. A rerun verifies what exists (by internal
  name and Swift marker), creates only what's missing, and stops with a diagnosis
  on any incompatible structure. It never drops, renames or rewrites a list that
  holds data.
- **Pre-checks before any write:** permissions (provisioning needs list
  management rights; daily use needs contribute rights), schema version and
  front-end compatibility.

**Testing strategy (product owner, 26 Oct 2026):** SharePoint has no container
image and no open-source implementation of its list protocol, so testing uses
three layers:

1. **Simulated SharePoint provider, every build, no cost.** A test-only provider
   on local Access tables that behaves like SharePoint Lists:
   - no transactions across lists;
   - relationships, validation rules and most unique constraints not enforced;
   - its own item IDs;
   - the 5,000-item view threshold and per-list index limits;
   - optional throttling;
   - injected failures: provisioning cut off at step N, a write dropped midway,
     a concurrent edit by another user.

   The same suite runs against the Local and Simulated providers in the extended
   tier. It covers idempotent saves, recovery, resumable provisioning,
   version and compatibility checks, conflicts and integrity checks. It can't
   cover Access's SharePoint connector or sign-in.
2. **A small commercial Microsoft 365 tenant for the connector.** Business Basic
   (trial or monthly), which runs the same SharePoint Online code as DoD365, apart
   from sovereign-cloud endpoints and the command's sign-in policies. It hosts the
   phase-1 research spike and a manual-tier runner against a dedicated test site.
   Synthetic data only.
3. **A DoD365 sandbox site for final acceptance.** A separate test site or
   subsite on the command's DoD365 with synthetic data only, requested from the
   site owner or ISSM. It is the only place to prove DoD365 sign-in and
   conditional-access policies.

Not chosen: a local SharePoint Server Subscription Edition lab. Its 180-day trial
needs Windows Server, SQL Server and a domain controller (16 GB+ RAM), and it
differs from SharePoint Online in sign-in and limits.

**Phasing:**
1. A research spike on a synthetic test site: linked-list behavior in Access,
   write atomicity and recovery, list limits, the provisioning API, sign-in.
   Record the findings and decisions here.
2. The storage-provider layer in local mode only, with no behavior change, plus
   the simulated SharePoint provider (testing layer 1). The full extended tier
   must pass unchanged.
3. The setup/settings screen (requirement 1), showing local mode.
4. SharePoint discovery, version and compatibility checks, and permission checks
   (requirements 2 and 4).
5. Guided, resumable provisioning (requirement 3), including interruption tests.
6. Shared-mode reads and writes with concurrency and conflict handling.
7. Moving data between modes through the production import, verified value by
   value.
8. Acceptance on the approved environment. Update the threat model, SV-1, SBOM,
   DevSecOps matrix, user guide, README and PROJECT_STATE.

**Acceptance criteria:**
- Local mode behaves exactly as today.
- Every listed feature works in both modes.
- Provisioning interrupted at any step resumes without data loss.
- An incompatible or newer workspace is refused with a clear diagnosis.
- Insufficient permissions are detected before any change is made.
- Two concurrent users can't silently overwrite each other's edits.
- Every earlier schema still imports.
- Native tests cover the provider layer in both modes; shared-mode tests run
  against a synthetic test site only.

## User review of 0.46.0-dev (25 Oct 2026)

Decisions from the user's manual review. Items marked *in development* are being built; the rest are open.

| Area | Decision |
| --- | --- |
| Counseling | Descriptive column headings; "Reason not performed" offers common reasons and accepts typed ones; no disposition date to enter; members without a grade/rate stay out of My Work; E-9 and CWO5 are not shown as due; a report created for a member starts with blocks 30–31 from the recorded session, never replacing entered values. *In development.* The pre-filled worksheet is not wanted for now. |
| Mark signed | Prompt and optional reason confirmed; dispose-of-drafts stays unticked; local approval not required. A group can mark all its reports signed; completing a group asks for one signed date and marks its reports signed; reopening a completed group reopens them. Reports list gets a Signed filter. *In development.* |
| Reports list | Rate column and rank filter. *In development.* |
| My Work | Contributors as a list; no Decision note in My review work (decisions are made in the report); review routes optional. *In development.* |
| Notes and replies | A running message log with author and time, a text box and Post; no conversation selector and no drafts. Existing replies and drafts stay in the data. *In development.* |
| Report editor | Boxes around grouped blocks (5, 10–13, 16–19); block 29 description across the available width; Mark signed drawn as a button. *In development.* |
| M01 | An unsigned `.accdb` opens on target workstations; signing is wanted for polish. Obtain a DoD PKI code-signing certificate through the command ISSM, sign the VBA project and verify normal-security launch. |
| Workspace location | Chosen and secured by each command in its approved PII location. |
| M07 | Keyboard order confirmed complete. Target the DoD Section 508 minimum (36 CFR 1194 / WCAG 2.0 AA) with a conformance report. |

## Low

| ID | Work | Scope and completion criteria |
| --- | --- | --- |
| L01 | Descriptive foreign-key names — implemented in development | Schema 32 uses descriptive names for all 34 formerly opaque major relationships. Production Settings import retains original fields/IDs and unchanged relationship enforcement/attributes in the fresh target; source schemas remain untouched. See the implementation scope below. |
| L02 | Clarify revision-ID field names — implemented in development | Writable schema-owned query aliases expose `ExpectedRevisionID` and `BaseRevisionID`; active batch/meeting/comparison consumers use them. Legacy physical columns remain the sole storage for exact historical import, without duplicated values or renumbering. See the implementation scope below. |
| L03 | Workspace size headroom — **proposed, not scheduled** | **User decision 27 Oct 2026:** not critical now. A workspace stops at Access's 2 GB limit. A synthetic 4,700-report workspace (10,340 saves) is 271.6 MB after compact-and-repair, about 26 KB per save, so the limit is about 78,000 saves: about 10 years for a command writing 1,300 reports a year at six saves each. Sizes there: `RevisionValue` 117 MB (65 MB of it rows that only record an empty field), `RevisionChange` 70 MB, `AuditEvent` 23 MB, everything else 62 MB.<br>**Planned, in order:**<br>(1) Settings & Help shows the file size against 2 GB and warns at about 1.5 GB. The guide explains starting a new workspace that carries over the directory, reporting seniors and latest RSCA baselines. No schema change.<br>(2) Stop logging empty-to-empty changes on a report's first save (measured saving 10%). No schema change; existing rows stay. The NAVFIT98A export guard in `modNAVFITAdapter` treats any logged field as edited, so this likely makes its native-comment protection apply to Swift-created reports. Confirm with the native-return test first.<br>(3) Sparse values (measured saving 24%; probably 10–20% on real reports, which fill more fields). Additive schema column `ReportRevision.ValueEncoding`: no value means a full copy, `1` means a missing field is empty. About 12 `src/` readers move to one accessor, and current-value lists keep the report cache. A per-revision field count and snapshot hash replaces missing-row corruption detection, which also delivers threat-model item 6. Import copies rows unchanged; old revisions are not converted. Draft disposal and recovery are unaffected. Test every persistence runner, schemas 1–35, mixed full and sparse revisions, disposal, restore and compare, then re-measure.<br>**Rejected after measurement:** delta storage (3% more than items 2 and 3, and draft disposal would have to rebuild signed revisions); one packed row per save (same saving as item 3, larger rewrite); altering existing columns for Unicode compression (hits the lock limit, then exceeds 2 GB). Only splitting workspaces or a server back end (M11) removes the ceiling. |

L01–L02 are implemented in **0.35.0-dev/schema 32** development source. See
[behavior, compatibility and verification](PROJECT_STATE_HISTORY.md#l01l02-relationship-names-and-revision-id-aliases).
No release or publication is included; historical fixture waivers remain in force.

**Remote follow-up, 0.37.0-dev/schema 33:** repaired the stale Members/workbench
fixtures without dropping their behavioral assertions, fixed non-grid proposal
projection and saved-export entry guards found by reordered task tests, and
added native label/focus/resize checks. M07 has development preparation and
targeted label/layout remediation, not accessibility acceptance. See
[current scope and verification](PROJECT_STATE_HISTORY.md#remote-regression-and-accessibility-preparation).

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

**M04/M05 follow-ups, 0.40.0-dev/schema 33:** workspace import now preserves typed
transaction outcomes and primary migration errors. Committed data is not reported
as rolled back, and committed/uncertain outcomes disable re-Apply with recovery
guidance. Four import/return/PDF view contracts require progress/status controls
and exact critical action bindings. Existing focused native tests cover confirmed
rollback/retry, an explicitly instrumented post-commit fault, immutable diagnostic
snapshots, source protection and wizard-control/binding corruption. No schema,
report geometry, release, signing or accessibility acceptance change is included.

## Initial source checks and partial delivery

**M02/M03 follow-ups, 0.44.0-dev/schema 33:** directory-batch report creation
invalidates earlier state before a failed new preview and requires the reviewing
contributor at Apply. Close/reopen, filter changes, failed Apply and VBA reset
require fresh review. Explicit transaction outcomes prevent false rollback claims
after a committed/display failure. The summary-group list replaces workflow UDF
lookups with a unique lifecycle join; full cell/order/filter equivalence, no-read
writes and alternating baseline/candidate measurements cover populated/empty,
missing lifecycle and inherited/blank/null closeout cases. Existing scalar write
guards and schema remain unchanged; these are bounded follow-ups, not release
or physical acceptance.

**M02/M04/M05 follow-ups, 0.43.0-dev/schema 33:** native-group import,
move-to-group and reviewed duplicate-merge previews require the same contributor
at Preview and Apply, and fresh review after failure/reopen/reset. Native group
import returns an explicit transaction outcome; postcommit display failures and
uncertain outcomes retain truthful feedback and disable duplicate Apply. Completed
moves/merges remain committed after display errors. Required control/action
contracts cover these three screens. Native tests reproduce the old cross-owner
behavior, then verify rejection, rollback, postcommit/display faults, a real VBA
reset and retained counts after reopening. No schema or release change is added.

**M04/M05 follow-ups, 0.41.0-dev/schema 33:** PDF import and NAVFIT return now
preserve typed commit/rollback outcomes through UI finalization. Committed/uncertain
states disable re-Apply with recovery guidance; confirmed rollback retains choices
for deliberate retry. Existing runners test transaction and display faults plus a
simulated uncertain outcome. Members/My Work contracts require critical controls
and exact action bindings, with native missing-control/wrong-binding checks. Report
geometry, schema and separate physical/security/accessibility acceptance remain unchanged.

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
