# Consolidation plan and final bounded review

21 October 2026. **Approved consolidation pass complete.** All nine workstreams
now have an implemented, quarantined or retained disposition. This closes the
bounded maintenance review, not every product backlog item or a proof of global
dead-code absence. No schema/runtime redesign or application release was made.
Current architecture is owned by [PROJECT_STATE.md](../PROJECT_STATE.md).

## Workstream outcomes

| Batch | Disposition | Result / retained boundary |
| --- | --- | --- |
| 1. Superseded release assemblers | Implemented | Three named assemblers deleted; historical release descriptions corrected. |
| 2. Temporary/build clutter | Reviewed and bounded | Six reproducible bytecode files deleted; four root helpers and 27 export folders held. Remaining mixed artifacts retained as evidence/unknown data, not assumed disposable. |
| 3. Browser prototype | Isolated | Five browser source files, builder and four browser tests held. Active fields.mjs stays byte-identical at its original native-input path. Browser storage/outputs untouched. |
| 4. Old assembly/packaging entry points | Implemented | Historical patch wrappers, duplicate preview builder, obsolete packaging/verifier and stale migration copy held. Current build/Demo/backup/upgrade paths retained. |
| 5. Tests and coverage | Reviewed and consolidated | Demo creation separated from feature suites; provenance self-contained; trait suite restored/modernized; message/snippet/folder/journal checks ported into current harnesses. Older contract-specific originals preserved. |
| 6. Documentation | Consolidated | Current contracts linked centrally; stale roadmap/recent versions corrected; misleading historical guide/navigation headers corrected. Unique notes/receipts retained. |
| 7. Repeated verification | Implemented | Reuse exact builder receipts; stale/unknown artifacts verify natively. Blank inventories and workflow/migration gates remain independent. |
| 8. Generator/runtime overlaps | Reviewed — retain | Active form layers, context hosts, mappings and distinct histories retained after producer/consumer review. No presumed dead public callbacks or schema branches removed. |
| 9. Dependencies/configuration | Reviewed — retain | Existing tools serve current workflows/clients; no justified dependency removal found and none added. |

Review-held content is not approved for final deletion. The user specifically
requested PROPOSED_PURGE, superseding the original preference against holding
folders. Holding is the chosen disposition for uncertain manual/historical use.

## Verified before/after counts

Counts are immediate files in maintained directories, not a claim about total
disk files; test runs create ignored artifacts and quarantine is relocation.

| Area | Before consolidation | Now | Explanation |
| --- | ---: | ---: | --- |
| scripts/ | 342 | 265 | Three deleted; 74 currently held; trait launcher restored. |
| tests/ | 126 | 118 | Four native modules and four browser tests held; trait harness restored. |
| prototype/ | 6 | 1 | Native field metadata retained; five browser files held. |
| src/ | 113 | 113 | All production VBA/catalogues/templates retained unchanged. |
| Dependencies / persistent utility scripts added | 0 | 0 | Existing entry points/harnesses modified directly. |

Holding area: **2,658 original files plus two inventories**: 91 historical
builders/helpers/tests/browser files and 2,567 generated VBA exports.
Permanent deletions: three assemblers plus six reproducible caches. The cache
deletions free 33,927 bytes; moving exports saves no disk space.

Initial Git inventory was 785 tracked paths (784 present) and 30 nonignored
untracked files, before the consolidation documents. The existing user deletion
of output/MANIFEST.json and unrelated workspace/reference changes were preserved.
Git counts are not a backup inventory for ignored binaries or user data.

## What changed and why

### Build and Demo workflows

Removed build-local-pdf-update, build-directory-merge-update and build-report-tools
release assemblers. The first required schema 26 despite current schema 27; the
manifest owns their current source/stages. Historical patch assemblers and old
packaging commands have no discovered current executable callers and are held,
not silently treated as safe to delete for external/manual users.

The canonical Demo builder is build-ux-admin-demo. It verifies a blank template
before creating a destination, seeds UX and directory history, restores normal
unload/startup behavior, compiles, runs populated-Demo/startup checks and preserves
the template hash. Existing destinations/populated templates are rejected.
Its old duplicate ran 14 normal pre-seed suites and an optional PDF suite.
Those 15 distinct suites remain available; their ownership is mapped in the
[Demo guide](UX_DEMO_BUILD_GUIDE.md). They were not declared equivalent or deleted.
Release verification remains separate and unchanged.

Artifact verification reuses only a complete adjacent builder receipt with
matching path/hash, app/schema, all current input/checker hashes, module coverage
and native completion flags. Missing/malformed/stale evidence triggers native
verification; ForceNative requests it explicitly. Current metadata and requested
blank inventory always run. No reusable independent-verification cache was added.

Single local runs on the same unchanged artifact: **2.1666 seconds with reuse
versus 87.2502 seconds forced native**. This is not a whole-build or cross-machine
speed guarantee. A populated synthetic artifact with a deliberately matching
structural receipt was still rejected as a Blank.

### Tests and preserved coverage

| Historical area | Current disposition / evidence |
| --- | --- |
| Trait persistence, invalid marks, partial failure, restore | Restored test-traits/TraitsHarness: 54 assertions, dynamic complete snapshot inventory, seven explicit trait keys, late-failure rollback, legacy missing-trait semantics and fresh-session invalid pending recovery. |
| Trait descriptions / fixed-page UI | Existing exact catalogue checker retained separately; fixed-page/report-tools suites own current UI. Retired checklist assertions were not ported. |
| Comment body/actor/parent guards | Current ReportToolsHarness expanded to 43 checks: exact 10,000-character body, contributor, blank/over-limit/cross-report/archived rejection, no rejected writes or revision changes. |
| Resolve/reopen standalone comments | Superseded current UI contract; original legacy harness held. Current replies remain independent of historical resolution; mappings/history remain preserved. |
| Provenance viewer | Existing suite now creates two synthetic receipts and raw values, including an unrelated sentinel; 24 checks and reopen pass without tools-results or TEST_DB. A populated input is rejected in its copy. |
| Native 128-column import capture | Not claimed by the synthetic viewer fixture; existing adapter/import/engine suites retained. The old row-zero fixture count of 128 became two explicitly seeded viewer fields. |
| Snippet pending/save and narrative integrity | Existing NarrativeHarness expanded to 46 checks, including long snippet text, type rejection, no automatic commit and exact saved text/history. |
| Folder batch failure/history | BuildCleanupHarness expanded to 31 checks, including injected late failure, original assignment/move rollback, failed journal and interrupted-operation explanation. |
| Library presets/snippet filters | Active UXReadRowsHarness and UXNavyWriterHarness retained after inspecting their current callers/assertions. |
| Native reconcile/undo | Active NAVFITGuard/PDFExpanded and other adapter suites retained. Historical source-specific tools cases remain evidence, not a claim of complete legacy-API retesting. |
| Group/grid comparison baselines | Pre-r126 GroupReadLegacy, UXGridLoadLegacy and My Work legacy helpers were quarantined after current-only repeatability/equivalence checks replaced their historical comparisons. Current group/grid/work-row tests remain active. |

The narrative launcher initially failed its old fixed 21-snapshot total after
the added snippet-save revision. It now checks after reopen that **each revision
has exactly one narrative snapshot**, and the complete expanded suite passes.
This replaces a fixture-size assumption with a stronger persistence invariant.

### Build clutter and data protection

Only pure source-check-GUID folders containing mod*.bas files were moved:
27 folders, 2,567 exports, 30,065,890 bytes. Original relative paths and every hash
were recorded before relocation in the
[export inventory](../PROPOSED_PURGE/SOURCE_EXPORTS_2026-10-21.json).
Inspected JSON/Markdown receipt metadata contained no exact-folder references;
possible external/manual diagnostic value remains the reason for holding them.

The remaining build root includes 290 databases, a lock file, 157 logs, 70 JSON
files and mixed diagnostic/helper artifacts. Their names do not establish ownership,
disposability or whether a workspace is populated. **Retain** these unknown/evidence
artifacts outside maintained source. No database was opened to classify it by
filename; no blanket build/output cleanup or old-output recreation was attempted.
Earlier releases, populated r116 workspace, diagnostic databases/logs, supplied
FORMS/NAVFIT materials and unrelated edits were preserved.

Deleted cache files had existing maintained Python sources. Only the six
individual files and then their empty directory were removed, without recursive
deletion. Pre-delete evidence:

| Cache under scripts/__pycache__/ | Bytes | SHA-256 |
| --- | ---: | --- |
| generate-pdf-map.cpython-312.pyc | 5880 | fd62bd9b4ad2ce276a6dc5f2244618d5efa2dc9cfc16a3962205f8814c7cc3f9 |
| generate-report-pages.cpython-312.pyc | 9838 | ef2a3d23dad625a47abc362a5cea6ef5e56db4537cec47ab1c6921fa553b6cd9 |
| make-pdf-expanded-fixtures.cpython-312.pyc | 7377 | b31087ea369230b94d080b452e17d7641e4b3112e1eb38f0db1e8751bf52e808 |
| make_pdf_sources.cpython-312.pyc | 356 | 65358773ea7507a547f62d6d08db2fabfa01c40d0df5e601ef7b7899bb252734 |
| pdf-structured-preview.cpython-312.pyc | 6767 | 6f2cc64d89e50f0f3adc1738a376a917e49b60b316b604b2f997407ecf55463c |
| sync-ux-help.cpython-312.pyc | 3709 | 0095e891908d98a45f4d592675e30ff8a96fe07dcf6961375bc5b9885dddae55 |

### Documentation and runtime audit

Header/reference triage covered all **127 planning Markdown files**: seven
current contracts/indexes and 120 checkpoint/service/evidence documents. Corrected
the roadmap's r127/C03/fresh-build claims, recent-development version and 20
misleading historical guide/checkpoint headers. Unique bodies, receipts and field
capacity/source details remain. This is not revalidation of every historical
statement; old notes no longer provide current instructions over PROJECT_STATE.

Audited all 72 manifest stages: one base-resource, 65 presentation, six final.
63 stage files contain managed-import guards. Reviewed the metadata/surface,
fixed-page, theme/visual, unload and final route owners. Repeated form access is
intentional layering; native event-stub deduplication preserves nonempty handlers.
No safe equivalent stage deletion was established. Shared dot-source context,
the required hidden group host and organizational-folder tools remain.

Reviewed legacy comment/note writers, ReportMessage conversion/drafts and migration
consumers. Legacy mappings, contributor/version ownership and separate domain
histories have distinct uses. Similar names do not establish redundant new writes
or authorize deleting retained history. Production migration and upgrade tools
still use canonical src/modMigration, not the quarantined schema-25 copy.
The separately authorized C05–C10 implementation follows the boundaries below.

## C08–C10 producer, consumer and navigation audit

Schema 29 adds a timeline index; it does not collapse histories that have
different retention, restoration or undo semantics. The current writer audit
includes SQL, DAO AddNew calls, generated events, macros and dispatcher aliases.

| Evidence | Active producers | Consumers and retention decision |
| --- | --- | --- |
| ReportRevision / RevisionValue / RevisionChange | CommitDraft and restore/import facades in modSwift | SavedFields, comparison, restore, review staleness and transfer verification. Retain snapshots, field changes and revision identity. One audit envelope per revision, not per field. |
| ReportEvent | ArchiveDraft; local approval/reopen in modRouteTemplates | Report state/review history and audit timeline. Retain distinct approval/archive evidence. |
| OperationJournal | BeginOperation / FinishOperation, called by import, batch, route, merge and undo services | Operations UI and guarded domain undo. Keep original operation author/start and mutable summary interface. Timeline records start and each outcome separately, including the actual completing actor. |
| FolderMove / FolderBatchChange | AssignFolder; bulk move/undo | Folder history, membership, batch undo and optimistic guards. Retain both individual move evidence and batch membership. |
| ReviewComment / ReviewCommentEvent | Import only; explicit compatibility conversion reads retained rows | Legacy mappings and resolution history remain importable. Current conversation posting does not write these tables or require resolution. |
| ReviewTask / ReviewTaskEvent | modRouting / modRouteTemplates | Assigned review steps, decisions, stale-revision checks and reopen. Retain these workflow histories independently of conversations. |
| ReportMessage | Unified tools; AddReviewComment and PostReportNote compatibility adapters; import conversion | Unified Notes and My Work conversation readers. Preserve author/time, legacy IDs, revision context and parent links. No duplicate ReportNote/ReviewComment posting. |
| ReportNoteDraft | SaveReportMessageDraft / retained private checkpoint facade | Contributor-owned mutable recovery with version guards. Retain the established storage name; posting creates ReportMessage and removes only the matching checkpoint transactionally. |
| ReportNote | Historical import backfill from pre-29 DraftNotes only | Original rows and legacy mapping remain. Normal UI reads no longer trigger conversion or write legacy notes. |

[modAudit](../src/modAudit.bas) supplies stable AuditGuid, source kind/version,
source foreign key, original actor/time, affected report/revision/task/folder and
explicit operation correlation. AuditActor retains the original profile ID/GUID;
its registration name/time are not presented as a historical event caption/time.
Historical nonmessage event captions remain unknown. Converted conversations
retain legacy note captions or the source profile's display label; that label
does not prove the person's name at the original event time. Timestamps
retain the application's existing local Access date semantics; no invented UTC
conversion occurs. Known operation GUIDs correlate batch/import/undo evidence;
otherwise events have independent correlation identities. Backfill infers links
only from unambiguous receipts or batch membership, not timestamp proximity.
Legacy operation outcomes remain in OperationJournal; backfill does not invent
their completing actor. Audit envelopes and domain changes commit/roll back
together. Existing domain undo creates new evidence; it never erases envelopes.

Conversation conversion allocates all legacy identities before linking parents,
checks missing/cross-report/cyclic ancestry and rolls back derived rows on failure.
Existing mappings are verified rather than silently repaired. New legacy API
results use negative canonical message tokens; positive IDs still identify
retained legacy comments. New callers should use AddReportMessage and positive
MessageID directly. Retired SetReviewThreadState rejects new resolution writes.
Approval and conversation posting remain independent. Modern schema-29 imports
retain DraftNotes without implicitly publishing it into the shared conversation.

### Generation and dispatch ownership

The six workspace destinations come from UXWorkspaceRoutes in
[modUX](../src/modUX.bas): Overview/frmHome, Reports/frmLibrary,
Summary groups/frmUXGroups, My Work/frmMyWork, Resources/frmUXResources and
Settings & Help/frmUXSettings. Rail generators and compact selectors consume
that list through scalar COM-safe facades. OpenScreen and UXNavigate share the
legacy tool dispatcher; report/group/senior pending-edit guards remain in place.

| Surface or route | Current owner / decision |
| --- | --- |
| frmHome / frmLibrary | Active overview and report collection; retained. |
| frmMajorWorkbench / frmBulkFolders | Required group context host and organizational folder tools; retained with distinct purposes. |
| frmReportTools / frmReportReviewPanel | Unified Notes/Review/Validation tools and current review panel. apply-ux-review-return creates the panel directly; apply-ux-my-work and apply-report-tools decorate it. |
| frmUXHelp / validation popout | Contextual help and validation remain modeless context-preserving tools. |
| frmHelp / frmChecklist / frmReviewComments | Removed duplicate base construction. Their public aliases route to current help or report tools. |
| frmUXReview | Removed temporary generation/copy on the managed fresh build. The standalone compatibility copier remains guarded for old disposable inputs. Alias opens current Review. |
| frmUXNoteComposer | Removed duplicate generation/decorations from note-entries/comments-ita. Alias opens unified Notes; private note recovery form remains. |
| AutoExec and generated callbacks | AutoExec invokes SwiftAppStartup. Form/control event expressions, source bindings and public dispatchers remain native verification inputs, not evidence of dead code. |

The existing finalizer still rejects unresolved retired bindings before deletion;
the fresh path no longer creates those five retired forms in earlier stages.
Other repeated passes remain distinct geometry, styling, context and unload
owners; no blanket stage deletion was justified. Fixed Front/Back pages, seven
Front Page traits and rail coordinates are unchanged.

The subsequent user-directed Members consolidation removes the roster-count
branch between frmMajorWorkbench and frmUXComparison. UXShowGroupMembers owns
new, empty, populated and last-member-removed presentation using the existing
comparison cache/grid and one toolbar in apply-ux-resource-tools. Legacy
workbench Members entry points forward to it; Review/History still use the
retained workbench. The canonical toolbar passes explicit group context into
shared-block and validation tools and seeds a hidden adapter only for legacy
member-selection/policy callers. Pending cells are drained first. Existing
domain services, private cache, revision writes and unload guards are reused.
The group-exit harness now checks creation, first/last membership transitions,
all retained actions, stale hidden context, policy/validation targeting and
report/review return paths; lifecycle checks target the canonical Members page.

[test-audit-conversations-navigation](../scripts/test-audit-conversations-navigation.ps1)
exports every actual form/macro as native text plus an inventory of destinations,
capturing control expressions, events and source bindings. Focused navigation,
context-help, My Work, group-exit, cleanup and fixed-page suites exercise the
supported entry/return paths and pending guards. The source/binding verifier
checks callbacks after compile/reopen. Physical/accessibility/printing acceptance
remains separate. Exact tested artifacts and schema gaps are in PROJECT_STATE.

## Final validation and evidence

This all-workstream pass passed **477 native assertions**:
54 traits, 264 fixed pages/buffers, 43 report tools, 31 cleanup/folder/journal,
39 reporting-senior/RSCA and 46 narratives/snippets, plus fresh-session trait
recovery and narrative reopen. **36 preflight checks** pass. Earlier focused
provenance, Demo, startup, catalogue and rejection checks remain recorded in their
receipts/implementation sections elsewhere. No obsolete broad historical runner
was used as the acceptance definition.

All **195 active build inputs** retain their pre-consolidation hashes, including
prototype/fields.mjs. The current development Blank remains SHA-256
55CC0C7C6B686A5F312FA52FC0DEA036D5FCA1F630B39F1ACDD39EAE266C8277.
Source preservation is also asserted by modified launchers/native suites.

Preflight receipt history retained from the earlier record:

- [20261021-122753-0615017c-beee-469b-a235-f727706964ba](../build/workflow/20261021-122753-0615017c-beee-469b-a235-f727706964ba/verification.json)
- [20261021-123758-6551392c-d7f9-4f89-a988-327995132171](../build/workflow/20261021-123758-6551392c-d7f9-4f89-a988-327995132171/verification.json)
- [20261021-124850-d19c3a12-ea57-4298-8e1d-24acbf4a3ac5](../build/workflow/20261021-124850-d19c3a12-ea57-4298-8e1d-24acbf4a3ac5/verification.json)
- [20261021-125319-16f18fde-5065-4c45-a16c-3e938641013a](../build/workflow/20261021-125319-16f18fde-5065-4c45-a16c-3e938641013a/verification.json)
- [20261021-125705-4b1f8a54-041c-4c34-9e27-4c7139cc4cd5](../build/workflow/20261021-125705-4b1f8a54-041c-4c34-9e27-4c7139cc4cd5/verification.json)
- [20261021-130632-e1b9f88b-a265-4ea6-94ca-1836ad18662d](../build/workflow/20261021-130632-e1b9f88b-a265-4ea6-94ca-1836ad18662d/verification.json)
- [20261021-130905-6490d93a-4216-4099-bbfb-66bfe1a34452](../build/workflow/20261021-130905-6490d93a-4216-4099-bbfb-66bfe1a34452/verification.json)
- [20261021-131351-e81c9d64-7dd5-47ea-8c96-8310973e10d2](../build/workflow/20261021-131351-e81c9d64-7dd5-47ea-8c96-8310973e10d2/verification.json)
- [Final all-workstream preflight](../build/workflow/20261021-133120-0271583f-5ac1-4b75-a3e0-2ce1a13b3c2b/verification.json)

## Remaining limits and next work

- Review-held files are not permanently purged; their inventories are the handoff.
- Schema 1–27 import remains implemented. r128's old-output testing waiver does
  not retire older support or prove current retesting. Future test policy remains
  a separate decision; no historical fixture matrix was run in this maintenance.
- No new full application generation, release upload or physical display,
  accessibility/printing acceptance is claimed. No production source changed.
- C04 policy applicability remains the next feature. Canonical field/schema work,
  audit/message retirement and UI-generation redesign remain separate backlog work.
- Static tracing cannot prove absence of external consumers or every dynamic
  callback. Retention/isolation is the explicit outcome where proof was lacking.

Authoritative structure: PROJECT_STATE for current architecture, AGENTS for
constraints, src for native sources, manifest/scripts for current build/checks,
tests for current behavior, docs for user/developer procedures, planning for
unique evidence/backlog, and protected output/build/reference areas. The sole
active file under prototype is native field metadata; PROPOSED_PURGE is not an
alternate executable project tree.

```text
PROJECT_STATE.md / AGENTS.md   current decisions and constraints
src/                         native VBA, reviewed metadata and templates
scripts/                     manifest build, Demo creation and focused checks
tests/                       current native behavior and test fixtures
prototype/fields.mjs         active native field catalogue only
docs/                        current user/developer procedures
planning/                    backlog and retained checkpoint evidence
output/ / build/             protected releases, workspaces and local test evidence
FORMS_REFERENCES_TEMPLATES/   supplied materials; NAVFIT98A folders also protected
PROPOSED_PURGE/              review-held originals and hash inventories
```

Final housekeeping: all 27 checked document link sets resolve; editor diagnostics
and diff whitespace checks pass (line-ending warnings only). Access process count
is zero. Existing unrelated edits remain local; no commit, push or upload was made.
