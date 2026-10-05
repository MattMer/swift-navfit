# Swift NAVFIT current-state review — 2026-10-04

Follow-up: the first three priorities below were implemented after this review.
The user changed report leaving to automatic save without a popup. See
[implementation and verification](UX_PRIORITY_FIXES_2026-10-04.md). This report
retains the original findings as a review snapshot. Priorities 4 and 5 were
subsequently addressed with compact viewport layouts and official-name identity;
see [implementation notes](UX_RESPONSIVE_IDENTITY_2026-10-04.md). Native visual,
DPI and assistive-technology acceptance remains pending; generated test logs may be
replaced by subsequent verification runs.

## Assessment

Swift NAVFIT has a substantial, tested drafting and workflow foundation. Its native redesign is still a development preview. The next milestone should be dependable, consistent daily workflows, rather than adding another major feature family.

The strongest areas are atomic report revisions, saved-field coverage, provenance, recovery services, guarded batch changes, and preservation during migration. The weakest areas are UI lifecycle protection, contextual navigation, small-window behavior, and consistency between individual reports and group tools. Passing service tests does not establish those user journeys.

This was a review, not a repair pass. Application source and existing user workspaces were not changed. Review fixtures were created separately under `build/`.

## Scope and fresh evidence

- Reviewed branch: `feature/ux-redesign`, HEAD `188c73935334d7e4c15fc348769f03790fc8f2f6`; latest application-source checkpoint `a1fe604`.
- Current source: **0.22.0-ux-preview, schema 21**. Prior distribution checkpoint: **0.21.0, schema 20**. These are different delivery states.
- Fresh build: `build/review-20261004.accdb`, SHA256 `1DBED77A7E0760D3B7BE03EF220978FF1F0BBD71616AA25A7FC1932A79AB446A`.
- All **24 foundation suites passed**, covering drafting, fields, traits, review, revisions, recovery, provenance, imports, major workflows and limited validation.
- **108 native editor/surface/dashboard assertions**, **33 populated-demo assertions**, blocked-content startup and enabled-content startup passed.
- **38 reporting-senior/RSCA/comparison assertions**, **6 workbench UI assertions**, **76 visible-control bounds checks**, and **425 navigation/content bounds checks** passed.
- Backup/restore passed closed-file waiting, quoted paths, byte hashes, no overwrite, restore and retained raw checkpoints.
- All four prototype test files passed: 41 model/field checks, 38 controller checks, 23 reporting-senior/calculation checks and actual ES-module report/group navigation.
- A synthetic schema-20 fixture seeded through the 44-assertion major-workflow harness migrated to schema 21. Every original row and field was compared; source hash remained unchanged; destination reopened. It included 2 reports, 11 revisions, 858 revision values, groups, batches, routes and comparison proposals. This is evidence for the tested fixture, not every possible operational workspace.
- A separate fresh demo was produced at `build/review-demo-20261004.accdb`, SHA256 `88135509898E3ABEAB6FF77B4E6DD51A98639A485F2BF13DCD1F20A3C233FCD1`. It is a review fixture, not a published replacement release.

Evidence is retained in `build/tests/foundations-suite-results.json`, the corresponding foundation logs, `build/tests/full-review-audit.txt`, `build/tests/review-*.log`, and `build/tests/review-schema21-migrated.accdb.migration.json`. Generated evidence and databases are excluded from Git.

## Highest-priority verified findings

### 1. P1 — Native document close can lose pending report edits

The custom navigation path uses `UXLeaveEditor`, but the generated editor has no `OnUnload` guard. Its unbound buffer is checkpointed every 30 seconds. Closing through the native document-close path before that checkpoint can discard values without the Save / Discard / Keep editing decision.

Reproduction on a scratch database: type `UNSAVED EXIT TEST` into the visible OfficialName control; confirm `EditorHasPending=True`; confirm empty `OnUnload` and zero checkpoints; close the document through `DoCmd.Close`; reopen/read the saved report. The saved name remained blank and recovery rows remained zero. This proves the bypassed document-close path; physical application-exit shortcuts were not separately tested.

Relevant source: `src/modUX.bas:48`, `scripts/build.ps1:322`, `scripts/apply-ux-workspace.ps1:106`.

Recommended change: one reentrancy-safe unload/exit policy used by both native and custom close paths. Cancel unload when the user keeps editing or save fails. Checkpoint deliberately before approved shutdown where appropriate; do not treat a timed checkpoint as a substitute for a leave decision. Regression-test closing before the first timer, invalid pending fields, save failure, and a value still being typed in the active control. Access provides a cancellable [Form.Unload event](https://learn.microsoft.com/en-us/office/vba/api/access.form.unload).

### 2. P1 — Bulk inputs can diverge from the preview being applied

`UXBulkChanged` updates pending values without invalidating `bulkBatch`. `UXApplyBulk` applies that stored batch. Reproduction: preview Ship/Station `PREVIEW A`, then change the input to `EDIT B`. Apply remained enabled and saved `PREVIEW A`.

The transaction service correctly applied the previously reviewed batch. The defect is the UI allowing edited inputs and an old apply target to coexist without a clear boundary.

Relevant source: `src/modUX.bas:405`, `src/modUX.bas:429`, `src/modUX.bas:439`.

Recommended change: invalidate preview and disable Apply whenever values, target membership or operation reason change. Require a fresh preview tied to those inputs. After applying, display a Results state and disable repeated application. Keep existing atomic commit and undo guards.

### 3. P2 — Report Review can inherit the wrong group context

The report Review action reuses `frmMajorWorkbench` and selects the report only if it is already present in that form's current list. An earlier group selection can leave that list restricted to another group's members.

Reproduction: open group A, then open a report outside A and use its Review action. The workbench retained group A and selected zero reports. The audit did not attempt a subsequent route mutation, so an incorrect route write is not claimed.

Relevant source: `src/modUX.bas:345`, `src/modUXDashboard.bas:72`.

Recommended change: open review explicitly with the current ReportID and saved revision. Clear unrelated group context and reset dependent selections. Make the target member/report period visible before every review action. Keep the existing Save and continue / Use saved version / Cancel decision.

### 4. P2 — Narrow resize leaves the old geometry in place

`UXResizeEditor` exits when width is below 12,000 twips or height below 7,500 twips, with no compact layout or fallback. After a large layout, a simulated 10,000 × 6,800 viewport returned False while the field panel still extended to 24,650 twips and the footer stayed at 12,200.

Relevant source: `src/modUXSurface.bas:286`, especially line 295.

The unsupported resize behavior is verified. The precise effect at each monitor resolution and Windows scale factor is a risk requiring native visual testing, not a measured result from this review.

Recommended change: collapse global navigation, shorten or collapse section navigation, and reflow content beneath a compact header. Provide safe vertical access when height is constrained and keep Save reachable. Apply the same viewport convention to group, comparison, senior and My Work screens. Test actual 1366×768 and 1920×1080 displays at 100%, 125% and 150% scaling.

### 5. P2 — The identity being edited may not match the header

The report header prefers local `MemberName` over official block-1 `OfficialName`. Reproduction: retain local display name `Legacy display`, save OfficialName as `UPDATED OFFICIAL NAME`, and reopen. The header still showed `Legacy display`.

Relevant source: `src/modUX.bas:139`, `src/modUXDashboard.bas:4`. Several group/work lists also select MemberName directly.

This is not a reason to silently synchronize independent saved names. Show the official name as the primary report identity, use a clearly identified local alias as supplementary text, and offer an explicit previewed composition action if users need to assemble a full name from parts. Use a consistent fallback for incomplete new reports.

## Source-confirmed UX gaps

These are implementation gaps or risks found in source, rather than additional runtime failures.

| Area | Current behavior / gap | Recommended direction |
|---|---|---|
| Shared report blocks | Bulk editor still uses a separate two-column, eight-fields-per-page layout; individual reports use compact block rows. | Reuse one section/block layout definition for reports and bulk editing. Show mixed values distinctly from blanks and deliberate clearing. |
| Shared defaults | Focused group Defaults still uses a field selector and a value input. | Show shared defaults in report sections, identify suggestion versus pending change, and preview explicit application to members. |
| Comparison | Compact trait list is read-only; editing requires selecting a member, changing separate mark controls and staging proposals. Headers include `T33`–`T39`. | Use a continuous subform with editable trait cells, official form-specific headings, locked-row explanations, and a separate staged-change preview. Preserve guarded proposal commits. |
| Review context | Several routes/actions still use the combined workbench and selectors such as `Report <id>` or `Group <id>`. | Focused report review with names, period, saved revision, current step and next action. Keep IDs in diagnostics. |
| Group calculations | Header shows PSG value/estimate but omits the calculation service's Partial flag and explanatory Reason. | Show included graded traits, blank/NOB exclusions, saved versus pending basis, historical baseline date, and why a value is unavailable. |
| Group report counts | Count includes all GroupMember rows; average excludes archived reports; PSG calculation rejects archived members. | Keep total membership visible, supplemented by active/archived/NOB counts and calculation eligibility. Explain the denominator. |
| Group lifecycle | Create/Add/Remove services exist; no usable rename/archive/recover group workflow was found. | Add group details, rename, archive/recover and deliberate member management. Preserve report data independently. |
| Guidance | Checklist still says to scroll below history or open old detail dialogs; some statements predate current rules/summaries. | Make findings jump to the correct section/block and rewrite guidance against current behavior. |
| Legacy identifier exposure | Normal report surface excludes DoDID, but workbench defaults iterate all saved keys and PDF import offers DoDID mapping. | Use SSN for current workflows; place legacy identifier handling in an explicitly labeled compatibility path. Retain stored data. |
| Return to Reports | Filters and selection have support; explicit scroll-position restoration was not found in native code. Prototype retains list scroll. | Verify and preserve the full search/filter/selection/scroll state after save, discard and return. |
| Focus | Section rendering moves focus to an anchor before hiding controls; resize also rerenders sections. | Preserve active control/caret on resize and logical focus after validation, dialogs and section switches. Verify with keyboard and a screen reader. |
| Performance | Header refresh can reread each member's seven trait values and scan the editor buffer after field updates. | Benchmark realistic 25/50/100-report groups before optimizing; cache or aggregate against saved revision/membership stamps if needed. No latency defect was measured. |

Sources: `scripts/apply-ux-layout.ps1:119`, `src/modUXDashboard.bas:97`, `src/modUX.bas:277`, `src/modMajorUI.bas:9`, `src/modRSDirectory.bas:131`, `src/modUXDashboard.bas:81`, `src/modUX.bas:196`, `src/modSummaryGroups.bas`, `src/modChecklist.bas:101`, `src/modWorkspace.bas:90`, and `src/modUXSurface.bas:200`.

## Feature inventory and remaining scope

“Implemented” here means present with tested behavior in the stated scope; it does not mean every screen has finished UX acceptance.

| Capability | Current state | Remaining work |
|---|---|---|
| Three report types / 78 saved fields | Implemented; section surface, SSN, checkbox flags, traits, narratives and name/contact fields | Final current-template caption/block review; consistent density; keyboard and narrow-window acceptance |
| Save / revisions / duplicate / restore / archive | Implemented atomic services; normal Save and history | Close-path protection; clearer save/lock/recovery state across every entry point |
| Recovery | Timed/manual checkpoints and guarded recovery implemented | Explain checkpoint age and unsaved exposure; native exit protection |
| Reports / folders / search / presets | Implemented with filters and saved views | Verify complete return-state preservation; consistent official-name display |
| Summary groups | Separate group list, counts, members, defaults, review, comparison and history | Group lifecycle, eligibility/exception clarity, shared report-form layout and compact screen behavior |
| Reporting seniors / RSCA | Reusable 16-field senior records, rank baselines, manual graded-trait count, optional exact sum and PSG calculation | Surface missing/partial reasons everywhere; clarify baseline provenance and reuse preview |
| Local review | Assignments, comments, routes, saved-revision decisions, stale guards and local approval locks | Contextual report review, guided stale-route restart, readable selectors and actionable lock reasons |
| Bulk / undo | Previewed atomic changes and guarded undo implemented | Preview invalidation, consistent report-block entry, clear mixed/clear semantics and Results state |
| Comparison | Automatic internal baseline capture, compact traits and staged proposal apply | Editable cells, clear pending group-average preview, consistent saved/pending state and historical comparison wording |
| Resources | Narrative snippets, review-route resources and senior directory | Consistent contextual insertion and return navigation |
| Imports / provenance | Read-only NAVFIT review, narrative import/reconciliation; supported PDFs map eight fields and retain other extracted values | Full supported-field mapping, matching/three-way returned-file reconciliation and target-app round-trip proof; unified Choose → Preview → Apply → Results |
| Validation | Local completeness, inline storage/entry validation and a limited versioned rule catalogue | Broader current-policy coverage and group eligibility; distinguish invalid, incomplete and limited Navy findings |
| Backup / migration / administration | Closed-file verified copies, contributor attribution, diagnostics, schema-20 transfer | Simplify operational wording and establish exact blank-prerelease verification |
| Printing / full NAVFIT export | Deferred or incomplete | Keep deferred until agreed; do not imply narratives fit the printed form |
| Network / authentication / synchronization / official signatures | Outside accepted redesign scope | No recommendation to add these now |

## UX recommendations

1. **Use the report as the primary working object.** Keep official block order and captions, member, period, save state, pending ITA and saved group average visible. Avoid requiring users to understand revisions, batches or internal snapshots before routine editing.
2. **Finish the form-filling model across all official sections.** Compact text rows; adjacent checkbox captions; dates, identifiers and grades sized to their content; wider narratives; consistent block boundaries. Preserve visible labels and useful focus targets as density increases. This follows [Fluent field guidance](https://fluent2.microsoft.design/components/web/react/core/field/usage/) on persistent labels, concise messages and actionable errors.
3. **Give every action a clear scope and state.** “Save draft” commits one report; “Stage changes” changes a pending comparison; “Apply reviewed changes” commits an approved preview. Show affected members and saved revision context before each mutation. Explain why locked/stale actions are unavailable.
4. **Make numbers auditable without making the screen technical.** Display counts and exclusions beside ITA, group average and PSG RSCA. Put formula details behind a labeled explanation. Identify an estimate from rounded historical RSCA, exact historical sums, and partial current-group data separately.
5. **Reduce navigation and decision load.** Report Review manages a report; group Comparison manages a group; Resources manages reusable records; Settings contains diagnostics. Contextual shortcuts return to the exact originating report/section.
6. **Provide usable empty and recovery states.** Explain what to create next, what a checkpoint contains, whether the saved baseline changed, and whether recovery restores pending edits or creates a new draft. Add clear first-run examples in seeded demos without placing synthetic content in blank operational builds.
7. **Make accessibility part of native acceptance.** Logical tab order, visible focus, persistent labels, text status, concise inline errors and predictable dialog return focus. Apply [Fluent accessibility guidance](https://fluent2.microsoft.design/accessibility) as design principles adapted to Access, then test the actual native implementation.

## Recommended implementation sequence

1. Repair native close/exit protection, bulk preview invalidation and report-review context. Add regression tests for the exact reproduced journeys.
2. Finish responsive navigation and safe small-window layouts. Verify focus preservation and the agreed resolution/scaling matrix in Access.
3. Share compact report-block layouts with bulk/defaults. Implement editable comparison cells with staged preview; retain backend safeguards.
4. Add group rename/archive/recover and transparent calculation context. Rewrite checklist navigation, labels and disabled-action explanations.
5. Consolidate current documentation and complete a separate blank prerelease with migration and exact-artifact verification.

## Release and documentation recommendations

`planning/CURRENT_STATE.md` is the clearest current entry point. `planning/ROADMAP.md:7` still calls 0.20/schema 19 current and lists several implemented foundations as future work. `planning/UX_REDESIGN_REVIEW.md` contains an earlier “keep schema 20” next-stage paragraph alongside later schema-21 justification. README's field inventory still describes detached panels. Historical material is useful, but needs explicit checkpoint labels rather than conflicting current instructions.

Use one maintained status matrix: implemented, partial, deferred and outside scope. Distinguish current source, latest seeded demo, and published blank build. Publish a blank prerelease only after workflow and native acceptance; retain separate seeded demos for feedback. Existing workspaces should transfer into blank targets, never into demos.

## Acceptance work still needed

- Real Access testing at 1366×768 and 1920×1080 with 100%, 125% and 150% Windows scaling.
- Keyboard-only creation, section changes, validation correction, save/reopen, group bulk editing, comparison, review, recovery and backup.
- Native screen-reader labels, error discoverability, focus transitions, and unavailable-action explanations.
- Physical native close/application exit, before and after a checkpoint, including failed save and keep-editing decisions.
- A realistic-volume report/group performance fixture.
- Final current-template field labels/block associations and remaining interoperability mappings.
- Access Runtime/x86 and another workstation, followed by exact blank-prerelease and representative migration verification.

No physical visual/DPI or screen-reader acceptance was performed in this review. Bounds assertions do not prove readable labels, absence of clipping/overlap, color contrast or prototype parity. Limited Navy checks and arithmetic tests do not establish complete official policy compliance. Printed layout remains deferred by the user's direction.

## Shared blocks and comparison follow-up

Review sequence step 3 is implemented in native Access. See
[shared block and comparison evidence](UX_BLOCK_TOOLS_2026-10-04.md) and the
[current checkpoint](CURRENT_STATE.md) for the final artifact and verification.
The original findings above remain a historical review.

## Group lifecycle follow-up

The rename/archive/recover part of sequence step 4 is implemented. See
[group lifecycle evidence](UX_GROUP_LIFECYCLE_2026-10-04.md). Calculation context
and checklist/navigation improvements remain subsequent work.

## Follow-up: calculation context and checklist (r20)

Completed the remaining two recommendations in sequence step 4: transparent
calculation context and section-aware checklist guidance. See
[implementation and verification notes](UX_CALCULATIONS_CHECKLIST_2026-10-04.md).
Existing arithmetic, saved transactions and schema 21 are retained. Current
artifacts and evidence are maintained in CURRENT_STATE.md. Manual display and
assistive-technology acceptance and prerelease/migration consolidation remain.

## Follow-up: documentation, candidate and guided imports (r21/r22)

Current workflows now have a native Settings & Help guide, matching written
guidance and a feature-destination table. Historical service guides point to
that current navigation. Reproducible candidate packaging verifies separate
blank/demo files, source modules, native guards, backup, startup and migration,
retaining logs and hashes. Schema-20 and schema-21 fixtures transfer with exact
workspace row/field checks; existing destinations, unknown tables and populated
templates are rejected. Candidates remain local review artifacts until manual
physical-display/keyboard/screen-reader acceptance; no release was published.

The next import UX recommendation is also implemented: contextual source
choice, full official-caption mapping preview, guarded Apply and retained
Results. Existing-source inspection and narrative reconciliation remain
available. Normal PDF imports use SSN, with explicit legacy-ID compatibility.
See [guided import evidence](UX_IMPORT_FLOW_2026-10-04.md) and CURRENT_STATE.md
for current artifacts, checks and remaining recommendations.

## r23 follow-up: focused review and Reports return

Implemented a dedicated saved-report review screen with named routes/reviewers,
saved revision/period, current step, next action, contributor/workflow guidance,
full step instructions and stale-route restart. Report-editor and group review
entry points open it. Atomic restart retains old decisions and creates fresh
pending steps; baseline, contributor, group/archive and approval guards remain.
Reports return restores selection and native vertical position by top report
identity after Access's close/focus processing, preserving filters/search/view.
The new scratch suite passed 41 assertions. Existing workflow checks passed 48
assertions plus both forced-shutdown cases. Native manual acceptance remains.

Final r23 follow-up evidence supersedes the preliminary counts above: 133 focused review/return/width assertions; 49 existing workflow assertions and two forced-shutdown cases; 44 major workflow assertions. All 49 standard source modules match final blank and seeded demo exports. Physical scaling and manual accessibility remain unverified.
