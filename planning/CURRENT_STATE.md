# Current state - 2026-10-04

Swift NAVFIT is a native Microsoft Access application for local EVAL, CHIEFEVAL
and FITREP drafting. The working branch is `feature/ux-redesign`, version
`0.22.0-ux-preview`, schema 21. The redesign is a development preview under
user review. The prior distribution checkpoint remains 0.21.0, schema 20.

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
- Summary-group list with visible current report counts; focused member,
  defaults, review, comparison and history views. Trait comparison captures
  its baseline automatically. Reporting-senior records and rank-specific
  historical RSCA/count baselines support group PSG RSCA calculations.
- Guarded startup: blocked Access content opens an instruction screen instead
  of invoking unavailable VBA. Office security settings are not changed by builds.

Existing drafting, library/folder operations, snippets, revision history,
local review/approval, bulk transactions and undo, imports and provenance,
backup/restore and migration services remain in the project. See the
[feature inventory](../README.md#current-features) and
[major foundations guide](MAJOR_FOUNDATIONS_GUIDE.md) for their supported scope.
Local approval is not an official signature; limited Navy checks are not a
complete official validation engine.

## Preview and verification

The latest local review artifact is
`output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r13.accdb`. It includes nine synthetic
reports, three groups, two contributors, a reporting senior with three RSCA
baselines, and review, approval, NOB, archive and recovery examples.
Earlier user workspaces and previews are preserved.

The latest checkpoint passed 108 native surface/workspace assertions and 33
populated-demo assertions, plus blocked-content and enabled-content startup
checks. Coverage includes resizing, pending-value preservation, compact trait
rows, Review/History positioning, readable generated captions, group counts,
validation and recovery. Earlier foundation regression results are documented
separately; they are not represented as a new full-suite run for this UX patch.

This repository tracks source, build/test scripts, the browser design prototype
and documentation. Access databases, local reference PDFs and user workspaces
are excluded from Git. This push does not publish a new Access binary release.
For a new synthetic review build, run:

```powershell
powershell.exe -NoProfile -File scripts/build-preview.ps1
```

Use `scripts/build.ps1` for a separate blank template. The browser prototype is
an interaction design reference, not a replacement application. See the
[demo guide](UX_DEMO_BUILD_GUIDE.md) and
[checkpoint details](UX_NATIVE_LAYOUT_CHECKPOINT.md).

## Remaining before redesign rollout

- Continue native visual review against the user's report-form density goals.
- Verify real laptop/desktop layouts and Windows 100%, 125% and 150% scaling;
  finish responsive navigation/collapse and remaining destination layouts.
- Manually verify keyboard navigation, focus and screen-reader behavior in
  Access; automated geometry checks do not establish accessibility.
- Finish contextual review/import/bulk UX and native close-window dirty guards.
- Complete current-template caption/block review, printed layout work and
  remaining interoperability/field mappings within their documented scope.
- Verify schema-20 transfer into a separate blank schema-21 redesigned build
  and complete exact-artifact prerelease checks before a replacement release.

Do not transfer operational data into a seeded demo. Use a separate blank
build and preserve existing workspaces during migration.
