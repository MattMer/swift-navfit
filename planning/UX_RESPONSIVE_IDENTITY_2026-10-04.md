# Responsive layout and report identity - 2026-10-04

Implements review priorities 4 and 5. Version 0.22.0-ux-preview and schema 21
remain unchanged. Earlier workspaces and r14 are preserved.

## Layout

The report editor no longer abandons resizing below 12,000 x 7,500 twips.
Compact mode replaces global and report-section rails with selectors, reflows
all official sections, resizes the field subform, and keeps the save area below
it. Very short windows use vertical document scrolling rather than leaving
stale wide geometry. Wide mode restores both rails. Trait guidance moves below
the marks when the surface cannot accommodate adjacent columns.

Reflow flushes raw active text into the existing unbound buffer and restores the
original visible field, caret and selection. It creates no report revision.
Validation and local approval/archival protection remain in effect.

Overview, Reports, groups, My Work, Resources, Settings, bulk, comparison,
reporting-senior directory/editor and group RSCA use shared collapsed navigation
and width-aware authored row geometry. Fonts remain at their existing sizes;
short windows use vertical scrolling. Mode changes restore authored geometry
before selecting the new mode, preventing repeated cumulative shrinking.

Access requires the editor module to be persisted after generating all controls,
then reopened before binding its unload procedure. The build follows that order
and verifies the persisted native save-on-close handler.

## Identity

OfficialName is the primary identity in the editor, report presets, Overview,
group members, member selectors, comparison, bulk preview and My Work. A
nonempty independent MemberName appears as Local name in the editor when it
differs. The fallback is MemberName, then Unnamed member. Comparison baselines
use the official name from their captured revision, preserving historical
context. Search continues to find either stored name.

No saved names are synchronized or overwritten by display formatting. Changing
structured name parts still does not silently compose a stored full name.

## Verification and limits

The dedicated native regression suite exercises narrow/short/wide reflow across
all report sections, active raw text/caret preservation, pending versus saved
revisions, official/local-name independence, identity in list destinations,
and horizontal control bounds across shared workspaces. Preview generation
runs this suite before seeding dummy data.

Actual monitor/DPI visual inspection, physical keyboard-only journeys and
screen-reader testing remain pending. Automated twip bounds do not establish
readable wrapping or native accessibility. Printing remains deferred.

## Exact artifacts and results

Seeded demo: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r15.accdb`.
It starts with nine synthetic reports, three groups, two contributors, one
reporting senior and three RSCA baselines, including review, approval, archive,
NOB and recovery examples. It is a local preview, not a GitHub binary release.

Build-time demo SHA256 (before user edits):
`E8BF486D628389CC458CB5ADD49B0A4A382B997B45FBDCBB1D05AD9457262DCE`.

Verified blank template: `build/ux-r15-blank.accdb`, SHA256
`F50A859CDFED7544C276142773EC8A134251073FEB7F67C717999F4053A05AAD`.
DAO checks confirmed zero rows in Report, ReportRevision, GroupMember,
SummaryGroup, RecoveryDraft, UserProfile, ReportingSeniorRecord and RSCABaseline.
Demo seeding left this source template unchanged.

- 2,069 responsive-layout, active-text/caret, identity and bounds assertions passed.
- 108 existing native surface, buffer, validation, caption and dashboard assertions passed.
- 51 close, bulk-preview and review-context assertions passed, plus both valid/invalid forced-shutdown recovery cases.
- 36 populated-demo assertions passed; normal blocked-content startup and enabled-content startup both passed.
- All 24 foundation suites passed on the final blank template.
- 426 shell navigation/content checks, 38 reporting-senior/RSCA checks, six workbench assertions and 76 visible-control bounds checks passed.
- Guided backup/restore passed closed-file waiting, quoted paths, hashes, no overwrite and retained raw checkpoints.

Evidence: `build/tests/ux-r15-preview.log`, `ux-r15-foundations.log`,
`ux-r15-shell.log`, `ux-r15-rs.log`, `ux-r15-major-ui.log`, `ux-r15-backup.log`
and the foundation-suite JSON/logs. Generated databases and evidence are ignored
by Git. The standalone shared-screen debug fixture was used to isolate Access's
retained inactive-form focus before the final complete suite passed.
