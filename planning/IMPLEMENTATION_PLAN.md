# Swift NAVFIT implementation plan

> Historical planning baseline: use [PROJECT_PLAN_STATUS_2026-10-04.md](PROJECT_PLAN_STATUS_2026-10-04.md) and [CURRENT_STATE.md](CURRENT_STATE.md) for current implementation and unresolved work. The original proposed designs and early environment findings below are retained for traceability.

Planning baseline: 3 October 2026. User authorized starting development after review. A limited local foundation prototype now exists; see `BUILD_STATUS.md`. No NAVFIT version, Access target, or PDF import adapter is certified compatible.

This document covers all nine requested planning deliverables. Proposed designs are decisions for review, not descriptions of an existing application. Unverified factual assumptions are prefixed **GUESS**. Evidence is in `planning/evidence/source-inventory.json`, extracted source packets and instruction text. Inspection tooling does not modify supplied PDFs.

## 1. Requirements and scope

Deliver one operational `.accdb` containing forms, queries, VBA, local report data, settings, policy guidance, comments, routing and revisions. Each new workspace starts without reports or personal information. Development exports and build tools may live separately. Backup files, NAVFIT interchange files and temporary export files are artifacts, not live backends. No required SharePoint, SQL Server, cloud or internet service.

Initial release is a sequential, single-user command workspace. One file is authoritative; independent copies do not synchronize. Swift prepares drafts and local approvals; NAVFIT performs final validation and official printing. Swift will not provide official signatures or submission.

Required first release: tested NAVFIT interchange; three distinct report models; coherent editor; organizational folders; policy-defined summary groups; inheritance and reviewed bulk application; versioned validation; committed revisions and comments; basic routing/My Work; comparison grid; staged import/export/reconciliation; blank distribution; backup, recovery and migration. Structured PDF import is a release objective with a separate feasibility gate; do not silently remove it if workstation restrictions prevent implementation. OCR, portable review package merge, advanced comparisons and comprehensive RSCA analytics may follow.

Additional useful features proposed from research and source inspection: offline policy bookmarks by block; missing-history checklist; summary-letter data completeness checklist (NAVFIT remains the printing destination); upcoming occasion reminders with effective policy dates; summary-group completeness attestation; a print-verification checklist attached to an export. These are proposals, not new Navy requirements.

## 2. Verified findings and unresolved evidence

### Local inspection

| Item | Verified observation | Limit |
|---|---|---|
| Workspace | Initially empty; four supplied PDFs appeared during inspection | No application source, sample NAVFIT databases or AGENTS.md found in workspace/checked parents |
| Development Office | Click-to-Run registry: `O365HomePremRetail`, x64, `16.0.20430.20092` | Does not establish Navy workstation configuration |
| Access binary | `C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE`, file/product `16.0.20430.20092`; COM class registered | License/startup/UI execution not proved |
| Access automation | COM activation failed with `0x80070520`, logon session unavailable | DAO execution, VBA compilation and Runtime behavior remain untested; failure is not proof Access cannot run interactively |
| NAVFIT installation | No NAVFIT directory found at top level of either Program Files directory; no running NAVFIT process observed | Limited search; installation elsewhere remains possible |
| Instruction | Supplied PDF has 174 pages; PDF page 1: BUPERSINST 1610.10H CH-2, 26 May 2026; page 2: CH-1, 16 Dec 2025 | Full validation catalogue and historical applicability review remain unfinished |
| EVAL PDF | XFA packet array; template 257,051 bytes; datasets 461 bytes; no canonical AcroForm fields returned; unencrypted | Blank template does not prove completed-form extraction or printed fit |
| FITREP PDF | XFA packet array; template 270,839 bytes; datasets 423 bytes; no canonical AcroForm fields returned; unencrypted | Same limitation |
| Chief PDF | Poppler: XFA, AES encrypted, copy/printing allowed, changes disallowed, JavaScript present | pypdf 6.10.0 fails with `Invalid padding bytes`; packet content and field mapping not yet extracted |

All source hashes are recorded in the inventory. XML wrapper fragments are intentionally not standalone valid XML; wrapper parse errors differ from failed template/datasets parsing. XFA scripts are data for inspection and must never execute during import.

### Authoritative external findings

Navy's current published page identifies NAVFIT98A `v2.2.0.332` / v33 Patch 2 and a separate conversion utility for v33-created databases; earlier database versions are outside that utility's stated support. This is published guidance, not evidence of the local installation or a tested interchange schema. [Navy performance evaluation software guidance](https://www.mynavyhr.navy.mil/Career-Management/Performance-Evaluation/eNAVFIT/).

Navy's forms page lists the EVAL, CHIEFEVAL and W2–O6 FITREP revisions as 05-2025. Do not extend W2–O6 support to other officer forms without a separate model and proof. [Navy Software & Forms](https://www.mynavyhr.navy.mil/Career-Management/Performance-Evaluation/Software-Forms/).

Microsoft documents an Access database size ceiling of 2 GB less system-object space. History and attachments therefore need measured capacity planning. [Access specifications](https://support.microsoft.com/en-us/office/access-specifications-0cf3c66f-9cf2-4e32-9568-98c1025bb47c).

Microsoft provides x86/x64 Access Runtime with Office installation constraints. Target approval and execution tests must precede choosing it as a supported configuration. [Runtime installation guidance](https://support.microsoft.com/en-us/access/download-and-install-microsoft-365-access-runtime).

Macro trust is a deployment prerequisite subject to organizational policy; do not change trust settings automatically or instruct users to globally enable macros. [Access trust guidance](https://support.microsoft.com/en-us/access/decide-whether-to-trust-a-database), [managed Trusted Locations](https://learn.microsoft.com/en-us/microsoft-365-apps/security/trusted-locations).

Research limitation: direct retrieval of several Navy pages/PDFs returned HTTP 403. Published page content was available through search; policy findings below are based on the actual supplied instruction, not search snippets.

### Evidence needed before dependency selection

1. Target Navy Windows version, Access product/build/bitness and full Access versus Runtime; permitted VBA/COM/ActiveX/native components; installation and approved trust process; printer/PDF driver.
2. NAVFIT Help/About software version and build date, executable path, architecture and installed supporting components. Obtain synthetic v332 databases and a genuinely blank database created by that installation. If legacy v33 support is wanted, obtain before/after-conversion pairs separately.
3. Completed synthetic PDFs saved by permitted viewers for all three forms, including multiline narratives, marks, NOB and empty values. Resolve encrypted Chief extraction using a suitable approved parser.
4. Command defaults and workflow preferences; expected report/history volume and retention direction from the command's records manager; authoritative RSCA baseline if that feature is required.

**GUESS:** Native Access forms, VBA and DAO will be permitted on intended workstations. **GUESS:** At least one approved interactive Access/NAVFIT test workstation will be available. **GUESS:** Completed PDFs retain structured datasets after saving. Each must be verified. No Swift application version or internal schema version has been assigned yet.

## 3. Proposed local schema and relationships

Keep NAVFIT tables outside Swift's internal model. The following table names and fields are proposed Swift designs only. Use AutoNumber local primary keys and Long Integer foreign keys; use stable GUID identifiers for workspace/report identity and future migration. GUID storage type will be selected during Access prototype validation. Date/Time values store UTC by convention with captured offset for display and audit interpretation. Short Text lengths are set in a data dictionary; narratives and structured snapshots use Long Text. Null and empty string remain distinct.

| Proposed tables | Essential fields and relationships |
|---|---|
| Workspace, AppMetadata, SchemaMigration | Workspace GUID/command name; app, schema, policy and adapter versions stored separately; ordered migration ledger |
| UserProfile, WorkspaceSession, Handoff | Profile GUID/display name; active session/heartbeat; from/to profile, transfer time and acknowledgment; not authentication |
| Folder | ParentFolderID self-FK, workspace FK, sort order, archived state; reject cycles |
| ReportingSenior, Member | Stable internal identities and selected administrative data; no SSN-based primary key; identity merges reviewed |
| Report | Report GUID, workspace/member/RS/folder FKs, form kind/revision, current revision FK, workflow/archived state |
| ReportAdmin, EvalDetail, ChiefEvalDetail, FitrepDetail | One admin record and exactly one appropriate form detail per report; draft nullable values; official block dictionary referenced |
| TraitDefinition, ReportTrait, NarrativeBlock | Form/revision-specific traits and block identifiers; unique report+trait or report+block; do not reuse Chief trait model blindly |
| SummaryGroup, GroupMembership, GroupException | Policy version, canonical eligibility key, RS identity; reviewed membership; explicit exception authority/decision evidence; membership changes revisioned |
| ComparisonSet, ComparisonMember, MeetingSnapshot, ProposedChange | Independent selection/order and frozen revision references; proposals do not change committed official fields |
| SettingDefinition, CommandSetting, FolderSetting, GroupSetting, ReportOverride | Typed setting keys and scope-specific unique keys; source/version for inherited values; official values committed separately |
| ChangeBatch, ReportRevision, RevisionFieldChange, RevisionGroupSnapshot | Author/time/reason/batch, parent revision, complete restorable official snapshot, typed old/new values with null flags, membership/context snapshot |
| CommentThread, CommentMessage | Explicit target: report/field/block/group-decision FK; replies reference parent message; category, resolved/reopened status and event history |
| RoutingTemplate, RoutingTemplateStep, RouteInstance, RouteStep, Approval, WorkflowEvent | Ordered profile assignments/due dates/return reason; approval ties to report revision AND group-context fingerprint |
| PolicyEdition, ValidationRule, ValidationRun, ValidationFinding, RuleTestCase | Versioned sources/predicates/severity and immutable run inputs/results; tests reference policy edition |
| ImportBatch, ImportItem, ImportValue, ExternalIdentity, ReconciliationDecision | Source hash/type/adapter/schema; raw staged value and proposed typed value; source database identity+external key; reviewed match/merge choices |
| ExportBatch, ExportReport, ExportGroup, ExportAttempt | Frozen revision/group manifest, adapter/template/software versions, file hash, verification/outcome and explicit attempt status |
| PdfMapping, NavfitAdapterManifest, FieldMapping | Form and revision-specific paths/transforms; verified external constraints; support-screen diagnostics |
| BackupRun, UpgradeRun, RecoveryEvent | Destination/hash/time/result and verification; source/target schema versions; no confidential values in diagnostic logs |

Relationships: Workspace 1:N folders/reports/profiles/batches; Report 1:N revisions, traits, narratives, comments and routing events; Folder 1:N child folders/reports; SummaryGroup N:M reports through membership; ComparisonSet N:M reports independently. Proposed current-membership uniqueness is report+grouping-context, refined after policy cases; never assume one group covers all historical contexts. Group snapshots refer to immutable report revisions. Approval and export references preserve those snapshots.

Enforce referential integrity and unique keys in Access where possible; enforce cross-table form-kind exclusivity, folder cycles and typed field constraints in the commit service. No cascade deletion of committed history. Archive instead of delete; retention-driven disposal is a separate reviewed process. Rule runs/cache calculations may be regenerated, but evidence-bearing approvals and exports cannot be silently purged.

Implement save through unbound forms or controlled buffers so official edits, revision records, audit changes, group effects and approval invalidation commit together. DAO transaction use is documented by Microsoft; crash/power-loss recovery still needs tests and backups. [DAO transactions](https://learn.microsoft.com/en-us/office/vba/access/concepts/data-access-objects/use-transactions-in-a-dao-recordset).

## 4. Screens and workflow specifications

| Screen | Behavior |
|---|---|
| First run / Home | Create workspace identity, command and local profile; show authoritative-file warning, file location, My Work, backup age and command name |
| Report library | Folder tree, search/filter, unassigned view, create/duplicate/import/archive/recover; moving folders changes organization only |
| Unified editor | Shared admin header; EVAL/Chief/FITREP-specific tabs with verified official labels and block numbers; keyboard tab order/shortcuts; draft save and visible timestamp; field guidance, inherited source, text capacity and findings |
| Summary groups | Explain eligibility key, exceptions and missing information; show member list, averages/capacity and completeness declaration; preview field changes before regrouping |
| Defaults and bulk preview | Precedence/source view; affected reports, old/new values, overrides, eligibility/findings and stale approvals; one reviewed change batch; locked reports skipped until explicitly reopened |
| Review / My Work | Assigned steps/due dates/returned reasons; revision diff, comments/replies/resolution; Accept/Return next step; local identity displayed |
| Rack and stack | Configurable compact grid with traits/ITA, proposed recommendation, duties/qualifications/achievements/readiness, status/issues; expandable narratives and side-by-side view; group-specific capacity; presentation mode masks unnecessary identifiers |
| Import / reconciliation wizard | Detect type/schema, stage, match, show field conflicts, confirm selected changes, commit recoverably; unsupported file produces useful explanation |
| Export wizard | Select reports/complete groups, validate snapshot and completeness, flag partial group export, review unresolved findings, write and verify NAVFIT file; status separates generated from NAVFIT-reimport/print-verified |
| Backup / upgrade / recovery | Guided closed-file backup, verification, restore to new path, migration preview and results |
| Support | App/schema/policy/adapter versions, Office/bitness, trust/runtime/component diagnostics and redacted batch failures; technical detail stays here |

Draft saves may be incomplete. Duplication creates a new report GUID and records provenance. Rollover previews selected admin values and clears evaluative fields unless individually selected. Restore creates a new revision using old values; it never rewrites history. Material official changes invalidate approvals; group membership or related member changes may also stale approval even if this report's fields stay unchanged. Comment-only changes normally do not invalidate official-field approval.

Workflow: Draft → In Review → Returned/Accepted → Ready for Reporting Senior → Approved for Export → Exported → Archived. Returned reports can re-enter review. Reopening a locked report creates an event and invalidates approval; editing an exported report preserves the old export manifest and creates a newer draft. Archiving can be reversed without removing history. Accepted means an internal route-step outcome, not an official signature.

Inheritance decision: command → parent folder → child folder → summary group → report override. Later parent changes NEVER silently change committed report values, including drafts. Recompute suggested defaults and mark differences; explicit Apply creates a reviewed change batch. Folder moves show newly applicable suggestions but do not commit them or regroup the report. Eligibility-defining defaults cannot silently determine their own group assignment; grouping always uses committed values and reviewed exception decisions.

Meetings store snapshots and proposals separately. Sorting ITA changes display/order only; it cannot assign a recommendation. A proposed trait edit previews ITA/group effects and must pass the ordinary commit/approval process. RSCA displays an authoritative dated baseline and coverage or 'Needs verification'; Swift-only partial history must not be labeled authoritative RSCA.

## 5. NAVFIT compatibility matrix and mapping strategy

| Candidate | Local software evidence | Sample schema evidence | Import/export/reimport/print | Support decision |
|---|---|---|---|---|
| v2.2.0.332 / v33 Patch 2 | Missing | Missing | Not run | First proof candidate, unsupported until tested |
| Earlier v33 build | Missing | Missing | Not run | Conversion research only; separate adapter/proof if requested |
| Pre-v33 | Missing | Missing | Not run | Reject; no claimed compatibility |
| Unknown/future schema | Unknown | Unknown | Not run | Reject or inspect read-only in Support; never guess an adapter |

Identify software version/build, database version indicator AND structural fingerprint. `.accdb` extension and Access engine version do not identify NAVFIT schema. Fingerprint includes table/column names, types/sizes/defaults/requiredness, indexes/keys/relations and adapter-specific version markers. Do not create guessed external table names.

Mapping record: adapter ID; NAVFIT software/build; template hash; schema fingerprint; form/revision; official block/label; Swift typed field; external table/column/key path; null/empty behavior; enum/code transforms; text/date/numeric constraints; derived versus entered; folder/group relations; preserve/recompute policy; test fixture and round-trip result. Field mapping for all three forms is explicitly NOT complete without sample databases. Preserve documented unmapped external values in staged provenance; if export cannot preserve a relevant value, reject that report/adapter rather than silently discard it.

### Exact compatibility proof procedure

1. On the authorized test workstation, record Access full/Runtime build+bitness, Windows, NAVFIT Help/About version/build, executable details, permitted components and printer. Start Access interactively; resolve the current session-specific COM failure only if automation is needed. Compile/run a disposable Access probe with DAO transaction/save/rollback, then a Runtime startup test if required.
2. In target NAVFIT, create a blank template and synthetic databases for EVAL, Chief and FITREP, multiple folders and multiple observed/NOB groups. Include punctuation, leading zeros, blank/null values, dates, every enum/trait, long/multiline narrative, addresses and boundary-length values. Save/export using the actual installed NAVFIT commands and record them; do not presume menu labels.
3. Hash originals and preserve read-only copies. Inspect scratch copies through DAO without running database macros. Enumerate all TableDefs (including applicable system/hidden metadata), fields/types/sizes/properties, indexes/key attributes, Relations, QueryDefs and database properties/version indicators. Record row counts and sanitized synthetic values; classify every table by purpose. Compare blank and populated databases.
4. Build the verified data dictionary, complete per-form field maps, recognized fingerprints and a template manifest. Test folders and group references, generated IDs, ordering, signatures/print artifacts and metadata; determine which values NAVFIT derives on reopening.
5. Create only a disposable minimal Swift proof harness after plan review. Import to staging; compare every relevant imported field/relationship. Export untouched content first using a copy of the verified blank template. Swift comments/routing/history tables never appear in export. Never modify original imports or use a populated database as a release template.
6. Repeat with one controlled administrative change, narrative change and trait change per form. Use frozen revisions and explicitly identified group impacts. Reject truncation, unknown codes and missing required related rows; do not export a falsely complete group.
7. Open/reimport the generated files in the exact target NAVFIT build, run its validation and print with the target printer/PDF driver. Inspect every official block, narrative wrapping/clipping/page flow, calculations and group summary. Character counts alone do not prove fit.
8. Re-export from NAVFIT, compare every relevant field and relationship to expected values. Maintain an explicit reviewed whitelist of NAVFIT-generated differences; no blanket exclusion of metadata. Test multiple reports/groups, NOB and applicable grouping exceptions. Record file hashes, schema, exact versions, diff, print output and reviewer result. Any unexplained difference fails compatibility.
9. Make controlled final-stage NAVFIT edits; reimport into Swift and prove reconciliation shows original frozen export, current Swift and returned NAVFIT values. Resolve conflicts explicitly; rerun relevant proof after changes to adapter/template/software.

Source files are opened read-only; imports stage into Swift. Exports are built at a unique temporary path, validated and closed before publishing to a user-selected new filename. A failed operation records a failed attempt and does not mark reports Exported. File creation is outside DAO atomicity: journal preparation, completion and recovery explicitly. Selected-report export preserves needed relational closure and original summary context; provide complete-group export as the safe default.

Duplicate detection: source file hash catches repeated batches; adapter+source identity+external PK catches same-source records; member/form/period/type/RS is only a candidate match. Swift GUID remains the identity. Reconciliation is three-way when a known export exists; otherwise treat it as a new staged import with explicit identity decisions. Null, empty, formatting normalization and deleted records require separate decisions. Confirmed imports create one identifiable batch; undo applies compensating revisions and refuses silent rollback over later edits.

## 6. Policy validation catalogue

Each rule stores RuleID, rule revision, form/revision, predicate and required inputs, source edition/page/paragraph/table/block, source hash, effective applicability and evidence, severity, message/correction, executable implementation reference, boundary/exception tests and review status. A ValidationRun freezes report revisions, group snapshot, rule edition and external-history completeness. Rules cannot claim historical applicability simply because a newer PDF contains them.

Severity: Error = confirmed applicable violation blocks export readiness; Warning = review concern; Needs verification = missing evidence/history/applicability; Local preference = command convention. Missing essential evidence may leave readiness unresolved without claiming a confirmed policy violation. Export readiness must not silently pass incomplete group capacity or unknown policy applicability. Narrative semantics are ordinarily warnings.

Source-verified catalogue seeds (not executable rules):

| ID | Scope/source in supplied instruction | Required cases before activation |
|---|---|---|
| GROUP-OFF | Encl (2), ch 1, table 1-3, printed 1-19/PDF 48; designator categories follow on 1-20–1-21 | Grade worn at end date, competitive designator grouping, duty/status, promotion status, ending date, type, billet, RS; dual-hatted exception; observed eligibility; complete designator category review |
| GROUP-ENL | Table 1-4, printed 1-22/PDF 51 | Current paygrade independent of rating; ACT/TAR grouping, INACT and AT/ADOS; UIC election/RS identity, promotion status, date/type/billet; observed eligibility |
| REC-TRAIT | Promotion recommendation instructions printed 1-16–1-17/PDF 45–46 | 1.0 trait, 2.0 trait, form-specific climate/accountability; O1/O2 non-LDO cap; prior recommendation history and withdrawal |
| REC-CAP | Printed 1-17–1-19, table 1-2 | EP and combined EP/MP limits by grade/designator; rounding, unused EP transfer, two-member exception, table sizes 1–30 and formula >30; verified complete observed group |
| NOB | Block 16 instructions printed 1-5/PDF 34 (CH-1); chapter 6 | Wholly versus partially observed, averages/marks/narratives, duration and occasions; separate form cases |
| CH2-OCCASION | CH-2 printed 3-1–3-2a, 3-6–3-7a | CMC/CMDCS/COB conditions, flag reporting relationship and dates; historical effective applicability and special billet election |
| RSCA-E56 | Paragraph 13-4g(5), printed 13-6/PDF 128 | Authoritative post-summary baseline and coverage; historical applicability review before activation |

Officer and enlisted group tables were visually inspected; extracted table text alone is insufficient for all designator exceptions. Table 1-3's observed-recommendation condition does NOT mean group reports by each recommendation value. Table 1-4's UIC exception needs a recorded election, not an automatic grouping choice. CH-2 changes must be reviewed for their exact applicability before code activation.

Additional catalogue coverage pending paragraph-by-paragraph review: required fields, dates/continuity, occasions/report types, adverse explanations, signatures versus draft state, export constraints, NOB nuances, retention and concurrent/operational reports. Treat technical adapter constraints separately from Navy policy. Form-specific test suites must include positive, negative, boundary, exception and missing-evidence cases; policy specialist review precedes activation.

## 7. PDF extraction feasibility plan

Confirmed XFA presence in all three supplied blanks. EVAL/FITREP templates contain repeated generic names as well as block-like names, so mapping must use full hierarchical paths/bind expressions and occurrence indexes, not leaf names. A blank datasets packet is not a completed-report fixture. Chief parser failure is a concrete feasibility risk; no Chief extraction claim.

Classification: verified XFA form/revision → XFA adapter; supported AcroForm → canonical field and widget agreement; flattened text → separate lower-confidence staging; scan → OCR-needed (deferred); unreadable/encrypted/unsupported → useful explanation. Unknown revision cannot reuse a known mapping automatically.

Feasibility work: inspect template/data bindings and checkbox/export values, signature fields and repeat structures; create completed synthetic fixtures with permitted viewer; extract datasets without evaluating JavaScript; compare every visible field to extracted values. Test compression, indirect streams, incremental updates, encrypted inputs and malformed/oversized files. Disable XML external entity/DTD resolution and impose resource limits.

**GUESS:** A self-contained VBA PDF stream reader plus approved Windows XML parsing could satisfy the no-install requirement for a restricted supported set. This is not a selected implementation. Python/Poppler used here are development tools and are not proposed workstation requirements. A general PDF parser involves compression/encryption/object resolution and cannot be replaced with string searches or an XML parser alone. Evaluate an approved embedded parser/library only after component approval and separate x86/x64 tests. If no feasible approved extractor exists, present the evidence and an explicit scope decision; do not require Acrobat automation by default.

Stage source form/revision/hash, full field path, raw/normalized value, confidence and conflicts. Confirm before creating/updating reports. Import provenance persists; editable drafts never inherit a valid signature merely because the PDF was signed. Store source hashes and relevant extracted data internally; storing original PDF blobs is optional pending capacity/retention decisions, not a required external live file dependency.

## 8. Distribution, migration, backup and recovery

Build from clean objects/schema and approved policy/configuration seed data. Never sanitize a populated operational file to create the release. Release QA queries every user-data table and inspects properties/embedded assets/hidden objects for reports, profiles, identifiers and confidential caches. Retain build manifest, hashes, object exports and synthetic tests separately. Operational format remains `.accdb`.

Dependency candidate is native Access/VBA/DAO with no optional ActiveX controls. This remains conditional on target testing. If Windows API calls are necessary, declarations/pointer types must be correct for x86/x64 and tested. [Microsoft 64-bit VBA guidance](https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/64-bit-visual-basic-for-applications-overview). Runtime needs explicit startup forms, navigation, dialogs, error handling and recovery paths, with no design-view dependency. Do not claim Runtime or older Access support from a successful full-Access run.

Maintain separate version identifiers for NAVFIT software, NAVFIT schema fingerprint, official form revision, policy edition, Swift app version and Swift schema version. Startup refuses newer unknown internal schemas with a helpful message. Missing/unsupported adapter disables exchange without blocking safe draft review where feasible.

Single-user enforcement proposal: prefer opening the authoritative workspace exclusively through the approved launch workflow. Microsoft documents exclusive DAO opens, but that does not prove a database can acquire exclusivity after its own startup; prototype actual application startup and bypass scenarios. [DAO OpenDatabase](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/dbengine-opendatabase-method-dao). Add session heartbeat/advisory blocking and explicit stale-session recovery. A table heartbeat alone is insufficient to guarantee exclusion. Explain residual bypass risk; local profiles and file internals are not strong authentication or tamper-resistant audit.

Handoff: save/close, record transfer, transfer the authoritative file through an approved local process, then acknowledge at next open. Do not copy while open or use synchronization folders for live use. Separate copies require deliberate replacement/reconciliation and are not concurrent workspaces.

Backup: save/close all connections before copying the entire file; use a permitted destination with matching sensitivity, verify hash and reopen a scratch copy to check identity/schema/key counts. A coordinated copy after Access exits is acceptable if an in-app closed-file helper is not permitted. Record request before close and verified outcome on next startup. Do not call a live byte-copy consistent. Restore to a NEW path, inspect and confirm authoritative replacement; never overwrite the only good file.

Upgrade: close old workspace, preserve verified backup, open new blank release and read old file read-only; detect supported source schema, preview counts/changes, migrate all reports, comments, replies, routing/templates/events, approvals, revisions, exports/imports/settings/provenance/identities. Use an explicit old-to-new PK map while retaining stable GUIDs. Target retains workspace identity, app version changes; partially migrated target is not authoritative. Validate counts, FKs, typed values, snapshots, histories and references; reopen in tested Access/Runtime before cutover. Failed upgrade leaves old file usable. Unsupported schema path rejects safely; repeat migration must detect duplicates or require a new blank target.

Archive is logical status with recovery. Monitor growth and test realistic histories well below the 2 GB ceiling. Prefer text provenance over obligatory PDF attachments. Choose warning thresholds after measurement; never purge revisions automatically to make room. Compact/repair only with exclusive closed-file control, a verified backup and post-operation checks. Retention/disposal follows approved command direction.

Protect file/backup/temp locations using approved OS permissions and workstation controls; minimize identifiers in presentation and diagnostics. Local encryption may be evaluated with deployment approval but does not itself establish CUI compliance. Expiry/cleanup of temporary exchange artifacts follows approved retention, with failed-cleanup paths reported. An `.accdb` user can inspect or alter data outside Swift; audit is operational history, not proof against tampering.

## 9. Phased implementation and acceptance gates

| Phase | Deliverable | Acceptance criteria / gate |
|---|---|---|
| P0 planning review and evidence | This plan, source inventory, target facts, policy/PDF evidence gaps | Review before full implementation; target configuration and required sample locations identified; unverified claims remain labeled |
| P1 compatibility proof | Synthetic NAVFIT fixtures, inspected dictionaries, complete per-form maps, blank template, minimal proof harness | Exact target Access+NAVFIT reimport/validation/printing passes for all claimed forms; every relevant value/relationship diff explained; source unchanged; unsupported schema rejected |
| P2 local foundation and core editor | Single-file schema, first run, save/duplicate/rollover/archive/restore, profiles, revisions, basic backup | All three models save incomplete drafts; restart preserves values; rollback failure injection leaves no partial revision; restored history creates new revision; backup/restore verified; single-use controls tested |
| P3 grouping, inheritance and validation | Separate folder/group/set models; defaults/bulk preview; activated reviewed rule catalogue | Folder moves never silently regroup/commit; all verified exception/boundary cases pass; overrides preserved; locked reports protected; group context changes stale approvals; missing history identified |
| P4 import/export/reconciliation | Staging, duplicate candidates, frozen snapshot export, recoverable batches | No silent overwrite/truncation; repeat import and three-way conflicts handled; snapshot stable after later edits; comments/history absent externally; failed publish not Exported; P1 regression rerun |
| P5 comments/routing/comparison | My Work, threaded comments, route templates, local handoffs, proposals and snapshots | Accept/Return/reopen audited; approvals revision/context-bound; sorting cannot assign recommendation; meeting proposals isolated; presentation masking verified; RSCA incomplete state explicit |
| P6 structured PDF import | Approved form/revision adapters and confirmation staging | Completed synthetic EVAL/Chief/FITREP values match visible fields, including multiline/mark/null cases; no scripts execute; signatures not transferred; malformed/unknown cases safe; component-free feasibility or reviewed scope change documented |
| P7 clean release/migration/pilot | Blank `.accdb`, deployment guide, tested migration, restore, support screen | Zero real data; all preserved categories reconcile across upgrade; power/failure recovery tested; each supported Access full/Runtime/bitness tested offline on relevant workstations; target NAVFIT round-trip/printing repeated from exact release |

P1 is a hard prerequisite to committing to a full editor implementation. P6 feasibility starts early in P0/P1 so a component restriction is discovered before release commitments. Phases can be adjusted after evidence, but compatibility, grouping correctness and safe data handling precede advanced features. No reliable duration estimate is possible until target access and fixtures are available.

Review decisions requested: confirm local single-file architecture and explicit defaults application; confirm initial form scope (E1–E6, E7–E9, W2–O6); identify exact NAVFIT candidate/Access target and permitted components; confirm workflow/retention priorities. The plan is ready for review; the application and compatibility proof are not complete.
