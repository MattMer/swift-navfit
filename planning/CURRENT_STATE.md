# Current state - 2026-10-04

Swift NAVFIT is a native Microsoft Access application for local EVAL, CHIEFEVAL
and FITREP drafting. The current source milestone is r24, version
`0.22.0-ux-preview`, schema 21. The redesign is a development preview under
user review. The prior distribution checkpoint remains 0.21.0, schema 20.

## Latest runnable build: r24

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
