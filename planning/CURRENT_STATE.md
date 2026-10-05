# Deployment version - 0.23.0-preview, schema 21

The deployment package is `output/Swift_NAVFIT_0.23.0-preview.zip`, containing
separate `Swift_NAVFIT_0.23.0-preview-blank.accdb` and
`Swift_NAVFIT_0.23.0-preview-DEMO.accdb`, migration tooling and guides. Open the
DEMO for synthetic examples. Existing workspaces need source-preserving transfer
into the blank build. See RELEASE_0.23.0_PREVIEW.md and the package MANIFEST.json
for verification and remaining acceptance work. Prior milestone filenames and
versions below are retained as historical evidence.

# Latest workspace update - r30, 2026-10-05

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r30.accdb` for the populated demo.
Blank: `build/ux-r30-blank.accdb`. Adaptive workspace layout now skips unchanged
sizes, and context refresh sizes only the active workspace. Redundant list
requeries were removed. Both builds compile; all 57 demo modules match source,
native event guards pass, and 1801 responsive assertions pass. See
WORKSPACE_PERFORMANCE_R30.md. Siltstone colors and existing feature scope remain.

# Latest presentation update - r29, 2026-10-05

The Siltstone palette is applied across all 51 native forms (1,717 controls).
Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r29.accdb` for the populated demo.
The separate blank is `build/ux-r29-final-blank.accdb`. Both compile; the demo's
57 installed modules match source and native event guards remain intact.
Version and schema remain 0.22.0-ux-preview / 21. See SILTSTONE_THEME_R29.md for
roles and verification limits. All r28 data and feature scope remain applicable.

# Current state - 2026-10-05

Swift NAVFIT is a native Microsoft Access application for local EVAL, CHIEFEVAL
and FITREP drafting. The current source milestone is r28, version
`0.22.0-ux-preview`, schema 21. The redesign is a development preview under
user review. The prior distribution checkpoint remains 0.21.0, schema 20.

## Latest runnable build: r28

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r28.accdb` for the populated
group-export preview. The separate blank template is `build/ux-r28-blank.accdb`.
Both compile and all 57 installed modules match current source. Summary-group
export previews saved members, computes promotion counts and produces one native
folder; exact PSG RSCA export is optional. Single-report export/returns remain.
Group returns and direct Swift printing are pending. Print exported reports using
NAVFIT98A's own engine after native validation. See
[group-export scope and checks](NAVFIT_GROUP_EXPORT_R28.md) and
[print-engine inspection](NAVFIT_PRINT_ENGINE_FINDINGS.md). Group native UI/print
acceptance remains unverified; prior r27 actual per-report round trips remain
the native evidence.

## Previous runnable build: r27

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r27-reviewed.accdb` for nine synthetic
reports, three groups, review work, reporting seniors/RSCA and recovery examples.
The separate blank template is `build/ux-r27-blank.accdb`.
Both artifacts compile and all 57 installed modules match current source. The
final demo passes synthetic-data and normal/enabled-content startup checks.
Exact hashes and native proof receipts are in
[r27 verification manifest](NAVFIT_R27_VERIFICATION.json).

All 74 supported saved report values plus type are mapped for EVAL, CHIEFEVAL
and FITREP. All 128 native report columns and native folder/summary content are
retained. Full supported-value import, single-report export and sealed-baseline
three-way returned-file reconciliation are implemented. Actual NAVFIT98A
2.2.0.332 edits were returned to Swift for each type; all 75 decoded values and
128 re-exported native report columns matched. Empty/null and long Unicode
preservation fixes are included. See [scope and proof](NAVFIT_EXCHANGE_R27.md).

Complete summary-group/batch export, legacy non-date counseling strings, changed
cross-file identities and other NAVFIT versions remain pending. Printed fit and
complete Navy readiness certification remain unverified. No GitHub release was
published; previous builds and populated workspaces remain preserved.

## Previous milestone: r26

The user's clarified priority #4 is implemented: expanded structured PDF import
catalogues (40 EVAL, 51 CHIEFEVAL, 36 FITREP saved fields), reviewed existing-draft
merges, stale/lock/recovery guards, source receipts and compensating undo. The
native extraction check and portable target-computer kit verify real XFA reads.
Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r26.accdb` for synthetic exploration;
`build/ux-r26-final.accdb` is the separate blank template. See
[PDF scope, evidence and workstation check](PDF_IMPORT_R26.md). Full field coverage,
other target-machine results, Runtime/x86 and physical/accessibility acceptance
remain outstanding. Earlier workspaces/builds remain preserved.

## Previous milestone: r25

The r25 rebuild adds live report diagnostics and an all-member summary-group issue list, with instruction, enclosure, page, paragraph/table and excerpts for verified Navy findings. See [scope and verification](UX_LIVE_VALIDATION_2026-10-04.md). Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r25.accdb` for synthetic exploration; `build/ux-r25-candidate.accdb` is the separate blank template. Full physical-display/accessibility acceptance and broader policy exceptions remain pending.

## Previous milestone: r24

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r24.accdb` for synthetic exploration. The separate empty workspace is `build/ux-r24-blank.accdb`. Version remains 0.22.0-ux-preview, schema 21; earlier builds and existing workspaces are preserved.

The three native-review priorities are implemented: safe report-list capture/reopening; one My Work inbox for assigned reviews, current route steps, ready local approvals and related unresolved threads; consistent heading/filter spacing, visible Save and lock feedback, readable list headings and supporting-screen navigation. Current contributors see actionable route steps rather than future steps. Stale assignment decisions are disabled. Opening work selects the actual route/thread. Resolution removes threads from the pending queue; history remains available.

Focused checks passed: 145 review/return assertions (including create, save, single-result search and reopening), 32 unified inbox assertions, 38 routing assertions and 28 comments assertions. All 50 installed standard modules match source and compile in both final artifacts; all 42 blank user-data tables are empty. Full preview verification and native visual results are recorded in `planning/UX_MY_WORK_2026-10-04.md`.

This remains a development preview. Full keyboard, screen-reader, actual Windows scaling, Runtime/x86 and release acceptance remain outstanding. No replacement GitHub release was published. The r22 candidate ZIP below remains an earlier packaged checkpoint.

## Latest UX work

- Persistent Overview, Reports, Summary groups, My Work, Resources and
  Settings & Help navigation, with command and contributor context.
- Section-based report editing with all 78 saved fields retained. Official
  captions vary by report type; structured name parts remain independent of
  stored full names. Normal editing uses SSN; legacy DoDID data is retained.
- Resizable report workspace with compact report block rows, adjacent checkbox
  captions, content-sized fields and closer label/input spacing throughout.
  Empty validation space is hidden; actual errors expand rows. Trait marks use
  compact rows with adjacent guidance. Review and History have focused layouts.
- Explicit Save draft, pending-change indicators, atomic revisions, recovery
  checkpoints, inline validation and a clickable error summary.
- Compact report layout replaces both navigation rails with selectors in narrow
  or short windows; larger windows restore the rails. Active text and caret
  survive reflow. Short windows retain vertical access to the action area.
  Group, comparison, senior and My Work destinations share viewport navigation.
- Report identities prefer the official block-1 name, with a labeled local name
  in the editor when different. Empty official names fall back to the local name,
  then Unnamed member. Saved names and name parts remain independent.
- Changed reports save automatically when leaving or closing, without a leave
  confirmation. Failed saves cancel normal report closing and preserve pending
  values. Forced Access shutdown retains raw recovery values if saving fails.
- Bulk changes require a current preview; typing, reason changes and changed
  group/report baselines invalidate Apply. Report Review resets unrelated group
  context and restricts the report and route selectors to the current report.
- Shared report blocks and group defaults use the same compact, official-block
  sections as report editing, with explicit mixed/blank/staged/clear states,
  full-value preview, atomic Apply and concurrent-default guards.
- Comparison has editable native trait cells, pending ITA and group averages,
  saved-order reload, auditable proposal replacement, before/after preview and
  stale/archive/approval guards. Temporary grid data detaches on native close.
- Summary groups support rename, reversible archive/recover, Active/Archived/All
  filters, status and retained report counts. Archived group workspaces explain
  recovery requirements; report data and group history remain intact.
- Summary-group list with visible current report counts; focused member,
  defaults, review, comparison and history views. Trait comparison captures
  its baseline automatically. Reporting-senior records and rank-specific
  historical RSCA/count baselines support group PSG RSCA calculations.
- Calculation details distinguish pending editor values from saved revisions,
  show sums/counts/exclusions and manual historical baseline provenance. PSG
  RSCA exposes exact/estimated, partial and unavailable reasons on initial load.
- Saved-draft checklist uses official captions, omits legacy/inapplicable inputs,
  and opens the relevant report section while preserving pending edits.
  Current guidance distinguishes incompleteness from limited Navy checks.
- A native workspace guide in Settings & Help covers current workflows,
  backed by a matching written guide and a complete feature-destination table.
- Contextual imports follow Choose/Preview/Apply/Results, with official mapping
  captions, full values, stale-preview guards and created-report navigation.
  SSN is the normal PDF identifier; legacy DoDID mapping requires opt-in.
- Guarded startup: blocked Access content opens an instruction screen instead
  of invoking unavailable VBA. Office security settings are not changed by builds.

Existing drafting, library/folder operations, snippets, revision history,
local review/approval, bulk transactions and undo, imports and provenance,
backup/restore and migration services remain in the project. See the
[feature inventory](../README.md#current-features) and
[major foundations guide](MAJOR_FOUNDATIONS_GUIDE.md) for their supported scope.
Local approval is not an official signature; limited Navy checks are not a
complete official validation engine.

## Earlier packaged r22 preview and verification

The earlier r22 local review artifact is
`output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r22.accdb`. It includes nine synthetic
reports, three groups, two contributors, a reporting senior with three RSCA
baselines, and review, approval, NOB, archive and recovery examples.
Earlier user workspaces and previews are preserved.

The r22 preview passes 30 native-guide/documentation consistency checks,
43 guided-import state/preview/result assertions, 28 native PDF assertions
and 12 read-only import-review checks. NAVFIT commit checks passed quarantine,
duplicates, rollback and all 2,472 retained provenance values. It also passes
23 calculation/checklist, 41 group-lifecycle, 90 shared-block/default/comparison,
1,790 responsive/focus/identity and 108 native editor-surface assertions.
It also passes 51 close/bulk/review assertions, two forced-shutdown scenarios,
36 populated-demo assertions and both startup modes.

The r20 checkpoint additionally passed 42 senior/RSCA, 54 integrated tools,
53 administration, 75 traits and 169 catalogue-fidelity checks. Those are prior
service evidence, not newly run r22 suites. All 24 foundation suites last passed
at r15; this is historical evidence. Exact packaged candidate tests are listed
in its MANIFEST.json, separate from these preview checks.

The blank is `build/ux-r22-blank.accdb`. The preview builder preserves that
source while seeding a separate demo with nine reports and three groups.
The candidate packager compares all source standard modules against installed
exports after normal VBA compiler formatting. It independently verifies blank
user-data tables, installed source, native unload handlers and migration.
The Member parameter prompt remains fixed by ordering the comparison query
by its full expression rather than its alias.

See [guided import changes](UX_IMPORT_FLOW_2026-10-04.md),
[workspace guide](UX_WORKSPACE_GUIDE.md),
[calculation and checklist changes](UX_CALCULATIONS_CHECKLIST_2026-10-04.md),
[group lifecycle evidence](UX_GROUP_LIFECYCLE_2026-10-04.md),
[shared blocks and comparison evidence](UX_BLOCK_TOOLS_2026-10-04.md),
[responsive layout and identity evidence](UX_RESPONSIVE_IDENTITY_2026-10-04.md)
and [earlier priority fixes](UX_PRIORITY_FIXES_2026-10-04.md).

Build-time SHA256:
- Demo: `1F172ABDF185DC939EA712A6E78246D3FAF2ED68E96DB69EDD29CB0D1F1AC7F1`
- Blank: `F958BA5A1A3488B05EADBFC1583DD2C7BA3DA49B243626EEB2937E6CF50DE4E1`

Opening or editing an Access file may change its hash. Full native DPI and
assistive-technology acceptance still remains.

This repository tracks source, build/test scripts, the browser design prototype
and documentation. Access databases, local reference PDFs and user workspaces
are excluded from Git. Local demos are separate from published binary releases.
For a new synthetic review build, run:

```powershell
powershell.exe -NoProfile -File scripts/build-preview.ps1
```

Use `scripts/build.ps1` for a separate blank template. The browser prototype is
an interaction design reference, not a replacement application. See the
[demo guide](UX_DEMO_BUILD_GUIDE.md) and
[checkpoint details](UX_NATIVE_LAYOUT_CHECKPOINT.md).

## Local candidate packaging

Use [the candidate guide](UX_CANDIDATE_GUIDE.md) for blank/demo distinctions,
source-preserving transfer and remaining acceptance work. Reproducible packaging:

```powershell
powershell.exe -NoProfile -File scripts/package-ux-candidate.ps1 `
  -TemplatePath build/ux-r22-blank.accdb `
  -DemoPath output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r22.accdb `
  -OutputDirectory output/candidates/Swift_NAVFIT_0.22.0-ux-preview-r22
```

The completed local package is
`output/candidates/Swift_NAVFIT_0.22.0-ux-preview-r22`, with a matching ZIP and
ZIP SHA256 sidecar. Its manifest records 42 empty user-data tables, all 46 source
standard modules matched to installed exports, native unload guards and 13
passing verification gates. Both schema-20 and schema-21 fixture transfers
passed exact workspace row/field comparison and reopen, with source/template
hashes unchanged. Unknown tables, populated templates and existing destinations
were rejected. The ZIP inventory and every packaged-file hash were verified.

The packager refuses existing destinations and publishes no GitHub release.
The distributed 0.21.0 checkpoint remains unchanged. Native manual acceptance
is required before this candidate becomes a replacement release.

## Remaining before redesign rollout

- Continue native visual review against the user's report-form density goals.
- Verify real laptop/desktop layouts and Windows 100%, 125% and 150% scaling;
  inspect responsive navigation, wrapping and remaining destination layouts.
- Manually verify keyboard navigation, focus and screen-reader behavior in
  Access; automated geometry checks do not establish accessibility.
- Focused review and Reports return are implemented in r23; complete manual
  review of the new screen and native return behavior.
- Complete current-template caption/block review, printed layout work and
  remaining interoperability/field mappings within their documented scope.
- Complete manual native acceptance before publishing a replacement release.
  Schema-20 transfer and exact-artifact checks are available in the separate
  local candidate package, with original workspaces preserved.

Do not transfer operational data into a seeded demo. Use a separate blank
build and preserve existing workspaces during migration.



Final r23 blank SHA256: `6C93E7F2F1F6D95A70FF62AC780770012FBCEBB06E88466434D7CAEAFEE59CB8`.
Final r23 demo SHA256: `46619E04F5A00F0B79A02569A15DA9465B4605944CFDA2741DB09D2CDD1992DA`.
