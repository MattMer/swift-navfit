# Swift NAVFIT project-plan status — 4 October 2026, r23

## r28 group export follow-up — 5 October 2026

Group export now previews saved members and writes one NAVFIT folder with
recalculated promotion totals, preserved native-only member values, a sealed
78-value-per-member baseline and optional exact PSG RSCA. It does not synthesize
summary-letter contact profiles. Group returns, actual group-native/print
acceptance and complete policy eligibility remain pending. Inspection found
readable managed print routines but no VBA COM print API; exported files can use
NAVFIT's own print actions. See `NAVFIT_GROUP_EXPORT_R28.md` and
`NAVFIT_PRINT_ENGINE_FINDINGS.md`. This supersedes r27's group-export limitation.

## r27 NAVFIT exchange follow-up — 5 October 2026

R27 supersedes the historical narrative-only import/export descriptions below.
All 74 supported saved values plus form type are mapped for each report kind;
all 128 native report columns and folder/summary source values are retained.
Single-report export and sealed-baseline three-way return handling are implemented.
EVAL, CHIEFEVAL and FITREP were opened, edited and saved in NAVFIT98A 2.2.0.332;
each return matched 75 decoded values and re-export matched all 128 native columns.
See [scope, native proof and remaining limits](NAVFIT_EXCHANGE_R27.md).
Complete group/batch export, changed cross-file identities, unsupported legacy
counseling strings, other NAVFIT versions and Runtime/x86 remain outstanding.

## r26 PDF-import follow-up

The user clarified the currently selected priority #4 as PDF mapping expansion,
existing-draft merge and target-computer extraction verification. R26 implements
40 EVAL / 51 CHIEFEVAL / 36 FITREP saved-field catalogues; guided fill-empty or
replace-value merges with full previews, revision/lock/recovery guards and undo;
native extraction diagnostics and a portable three-form workstation checker.
All three paths pass locally on Access 64-bit/Acrobat. Other target computers must
run the checker before their configuration is called verified. Full PDF mapping,
unnamed checkbox proof, Runtime/x86 and print appearance remain outstanding.
See `PDF_IMPORT_R26.md` and `CURRENT_STATE.md`. The original r23 tables and
recommended-focus numbering below are historical context, not this task's
priority definition. R25's live source-backed validation also supersedes the
older validation-coverage description, within its documented limits.

## r24 follow-up

The native-review top three priorities are implemented in r24: report-list reopening, the unified My Work inbox, and shared layout/state feedback. My Work now includes simple assignments, current ordered-route steps, ready local approvals and related unresolved comments, with named reports/contributors, next actions and stale/locked assignment guards. Future steps stay out of the current queue; resolved threads leave pending work. Existing tables support this without a schema change. Route/comment due dates and comment assignees are deliberately not invented.

See `UX_MY_WORK_2026-10-04.md` and `CURRENT_STATE.md` for r24 artifacts and verification. The r23 evidence and table below remain the historical checkpoint; their statement that My Work only queries ReviewTask is superseded. Next focus: broader native mutation/keyboard/DPI acceptance, supporting-screen density and tab accumulation, followed by the documented interoperability milestone. Printing remains deferred.

This is the current status of the original application plan and the accepted
UI/UX redesign, including subsequent user changes. Historical plans and build
notes describe earlier checkpoints; they are not current completion claims.

The runnable development preview is 0.22.0-ux-preview, schema 21:
- Synthetic demo: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r23.accdb`.
- Empty workspace: `build/ux-r23-blank.accdb` (42 user-data tables verified empty).
- Earlier local packaged candidate: r22. The redesigned app has not replaced
  the prior 0.21.0/schema-20 distribution.

## Original application plan

| Plan area | Implemented / verified scope | Outstanding work |
|---|---|---|
| Local architecture and setup | Native single-file Access workspace; local profiles, command identity, settings and data; blank and separate synthetic builds. No required live backend/network. | Approved workstation dependencies, realistic volume/capacity testing and operational handoff/session controls. Profiles remain attribution rather than authentication. |
| 1. NAVFIT compatibility proof | Inspected supplied Patch-2 sample/blank and installed executable evidence. All-field raw staging/provenance comparisons exist. User confirmed Comments in Chief block 40, EVAL block 43, FITREP block 41, with reports opening successfully. | Complete verified per-form field/code/folder mapping; controlled export, reimport and all-value round trip. Opening three narrative probes is not full compatibility certification. Older/future versions remain unsupported. Printing checks remain deferred by the user. |
| 2. Report editor | EVAL, CHIEFEVAL and FITREP; all 78 current saved fields, official sections, SSN, checkbox blocks, trait guidance, narratives/snippets, independent name parts, incomplete saves, duplication, archive/recover, ITA/group averages and inline errors. | Final current-template label/block/capacity review; additional unmapped fields such as qualifications/achievements and RRS meanings; deliberate rollover with selected administrative values and evaluative clearing; native keyboard/visual acceptance. Stored field count does not establish complete official-form coverage. |
| 3. Folders, groups and comparison sets | Independent folder hierarchy, explicit summary groups and comparison baselines. Group list/counts, membership, rename/archive/recover and preserved history. | Full policy-defined grouping keys, eligibility/exception rules and complete-group evidence; preview of membership/average/recommendation consequences when grouping fields change. Explicit local membership is not policy certification. |
| 4. Defaults and bulk changes | Command/folder ancestry/group/report override precedence as suggestions. Shared blocks/defaults use report sections; mixed/clear/staged states, full value preview, atomic batches, stale guards, approval locks and guarded undo. | Integrate full grouping/validation/capacity consequences into previews; finish visual/source-of-default clarity across all contexts. Parent changes do not silently rewrite committed reports. |
| 5. Policy validation | Entry/storage validation, saved-draft completeness, versioned source-linked rule-pack foundation with five Navy draft checks, recorded findings and tested promotion-capacity arithmetic. | Broader form-specific required/date/occasion/type/NOB/adverse/continuity/export rules, full grouping/EP/MP exceptions, complete severity/corrective guidance and historical applicability. Recorded rule-pack coverage is bounded; extending it requires source verification. No complete Navy/export-readiness engine exists. |
| 6. Changes, comments and revisions | Stable local report identities; atomic revisions with field snapshots/diffs; restore as a new revision; report/revision comment threads, replies and resolve/reopen history; operation/batch provenance. | Explicit field/block/group-decision comment targets, structured categories, trait justification and breakout rationale as planned. Current comment service is report/revision scoped, not every proposed target. |
| 7. Routing and local handoffs | Local assignments/due dates/My Work, ordered templates, saved-revision routes, contributor decisions, stale guards, local approval/reopen. r23 adds focused named review, next-step guidance and atomic stale restart with old decisions retained. | Unify simple ReviewTask assignments with RouteRunStep work in My Work; route-step due dates and unresolved-comment queue. Current My Work queries ReviewTask, not the complete ordered-route queue. Explicit handoff acknowledgment/session controls and export-dependent workflow stages remain. |
| 8. Rack and stack | Compact editable trait cells, ITA/group averages, staged recommendation/trait proposals, automatic comparison baseline, before/after preview, atomic Apply, stale guards; underlying meeting ordering/snapshots. Reusable seniors and manual rank RSCA/count/sum baselines support exact/estimated PSG calculation. | Configurable comparison columns, selected duties/qualifications/achievements/readiness, issue counts, expanded/side-by-side narratives, presentation mode, structured breakout rationale and complete-group recommendation capacity in the meeting workflow. Do not automatically assign recommendations from ITA sorting. |
| 9. PDF import | Supplied May-2025 interactive XFA templates for all three forms; eight-field partial mapping; synthetic filled examples; native Acrobat extraction/Access tests; Choose/Preview/Apply/Results, source receipts/XML and duplicate guards. | Full supported-field mapping, merge into existing reports and approved target-workstation extraction dependency. Current adapter requires installed Adobe Acrobat automation. General AcroForms/flattened PDFs/scans are not supported import paths; OCR remains a later enhancement. PDF appearance/printed fit is unverified. |
| 10. Import/export/reconciliation | Read-only NAVFIT inspection, narrative-only import, source provenance/quarantine, source-hash duplicate checks, explicit two-way narrative reconciliation and compensating undo. Guided mapping scope/full values/results. | Complete report import; stable cross-file identity matching; full conflict handling and three-way returned-export reconciliation. NAVFIT export, frozen export manifests/outcomes and full target-app round-trip proof remain unimplemented. Altered-file duplicate detection is not solved by source-hash guards. |
| 11. Distribution, updates and recovery | First run, workspace identity, guarded startup, automatic save on leaving, pending checkpoints/crash recovery, verified closed-file backup/restore, schema/version checks and source-preserving migration scripts. r22 packaged schema-20 and schema-21 transfer fixtures passed exact comparisons. r23 blank/demo are separate and audited. | Final redesigned-release packaging/acceptance and current comprehensive regression run; Access Runtime/x86 and approved workstation/trust/component tests; realistic retention/volume and failure-recovery drills. Earlier migration fixtures are evidence, not a universal guarantee for every workspace. |

## Accepted UI/UX redesign

| Plan step | Current status | Remaining |
|---|---|---|
| Clickable prototype | Built and reviewed iteratively with the user; three report types, group workflow, comparison/import/bulk and edge states informed native work. | Retain it as the design reference; verify all final native journeys against it rather than treating prototype accessibility as Access evidence. |
| Journey review | User feedback incorporated: current block labels, SSN, checkboxes, ITA/group average, report-form bulk editing, editable condensed traits, independent group list, automatic baseline and senior/RSCA directory. | First-time-user and experienced-user end-to-end native acceptance. |
| Shared UI conventions | Native Segoe UI, light/blue hierarchy, shared navigation and compact report sections, text statuses, explicit actions, errors and responsive focus-preserving layout. | Visual consistency, wrapping, logical focus/tab order and disabled-action explanation review across every destination, including legacy contextual tools. |
| Native form rebuild | Core Overview/Reports/editor/groups/comparison/My Work/seniors/resources/settings/import/help surfaces implemented; unbound buffers and backend services preserved. r23 adds focused review and native Reports viewport return. | Physical 1366×768/1920×1080 checks at 100/125/150% scaling; keyboard and screen-reader checks; any remaining density/context polish uncovered by these checks. |
| Legacy entry points | Editor detail controls are reflowed into official sections; feature destinations documented; contextual groups/review/bulk/import tools added. Legacy service/workbench data remains. | Complete manual journey inventory to ensure every entry point consistently returns to the right context; remove or simplify residual technical wording without dropping capability. |
| Rollout | Blank and seeded preview artifacts, guides, candidate packager, installed-source audit and migration verification available. Schema 21 has the justified senior/RSCA extension; r23 adds no schema change. | Publish a replacement prerelease only after native acceptance and exact release verification. The previous distribution/workspaces remain preserved. |

The original Save-and-leave prompt was superseded by the user's later direction:
valid changes save automatically on leaving. Save failures preserve edits and
block ordinary closing; invalid forced-shutdown edits remain recoverable.
Review with pending values still offers Save and continue / Use saved version /
Cancel because decisions operate on saved revisions. Manual roster freezing was
also removed as requested; comparison captures its baseline internally.

## Acceptance status

Implemented and exercised in scoped tests: all current saved fields retained;
section transitions preserve pending values; ordinary saves create atomic
revisions; default save/discard/failed-save/recovery guards; explicit saved-version
review; preview invalidation/atomic Apply/undo; scoped group/report actions;
native Reports return; separate blank/demo and source-preserving migration.

Still unverified manually: a first-time user can create/save/reopen unaided;
experienced users reliably find each official block; every journey works fully
by keyboard and screen reader; physical display/scaling layouts; complete
native visual consistency and operational-workstation compatibility.

Fresh r23 evidence: 133 focused review/return/width assertions, 49 existing
workflow assertions and two forced-shutdown cases, 44 major workflow assertions,
36 final seeded-demo assertions and both startup modes. Both final binaries
compile and match all 49 standard source modules; review/resize/return hooks are
installed. The build's broader import/help/calculation/group/comparison/editor
checks passed before the final focused-screen sizing/lock refresh patch.
All 24 foundation suites last ran at r15; they were not all rerun against r23.
Physical scaling and accessibility are not established by automated geometry.

## Recommended focus

1. **Close native UX acceptance and route-work discovery.** Test the demo's real
   create/edit/save/reopen, group bulk/comparison, stale restart and recovery
   journeys on actual displays/keyboard. Next contained implementation: bring
   ordered route steps into My Work with an actionable next step and unresolved
   comments, then polish any remaining legacy wording/context. This addresses a
   verified queue gap rather than adding another disconnected tool.
2. **Complete NAVFIT interoperability as the major implementation milestone.**
   Build a tested per-form mapping matrix from controlled synthetic target-app
   fixtures, then complete staged import, template-based snapshot export and
   returned-file identity/conflict handling. Prove every mapped value in NAVFIT.
   Keep printing deferred as requested, and label any proof that omits it.
3. **Finish policy-defined groups and validation.** Tie confirmed grouping keys,
   exceptions/completeness and recommendation limits to membership, edits and
   preview outcomes; expand versioned rules with source/effective-date and
   boundary/exception tests. This is necessary before official readiness claims.
4. **Verify and package the redesigned release.** Refresh the full regression
   set against the exact release artifact, verify populated migration, test
   Runtime/x86/approved workstation configurations, then distribute a clean
   blank prerelease alongside a clearly labeled synthetic demo.
5. **Later enhancements.** Deliberate rollover, targeted/group comments,
   richer rack-and-stack/presentation and full PDF mapping. OCR, portable review
   packages, broader historical analytics and optional reminders remain later
   work. Printing stays deferred; network integration, synchronization,
   authentication and official signatures remain outside the accepted scope.

The nine original planning deliverables are documented in IMPLEMENTATION_PLAN.md
(requirements, findings, schema, screens, mapping strategy, rule structure, PDF
feasibility, distribution/recovery and phases). Some evidence there is explicitly
historical. This status report and CURRENT_STATE.md describe implementation now;
documented designs must not be mistaken for completed features.

## Source control and collaboration

The local repository is on `feature/ux-redesign`; origin is the configured
MattMer/swift-navfit GitHub repository. Last local commit is `188c739` (current
UX-preview-state notes). There are accumulated uncommitted/untracked changes,
including r23 work. This turn did not commit, push or publish a replacement
GitHub release. The remote branch was not fetched or compared, so remote
freshness is not asserted. Keep source/test/documentation review separate from
native binary distribution; generated databases remain local/release artifacts.
