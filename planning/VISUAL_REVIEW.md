# Visual review and semantic colors

2026-10-19. Target: full desktop Microsoft Access. Free Access Runtime support is not required.

## Color roles

| Role | Color | Meaning |
| --- | --- | --- |
| Navigation rail | Navy #142D4E | Main workspace navigation |
| Current workspace/page | Teal #0E706B | Active location, also bold |
| Primary action | Blue #245B8F | Save, open, create, import or current next action |
| Help | Purple tint #EEE8FA | Context help and workspace guide |
| Attention | Amber tint #FFF0CF | Unsaved work, scope warnings, outstanding issues |
| Error/destructive action | Red tint #FCE8E8 | Failed save, invalid state, archive/discard/return |
| Saved/clean | Green tint #E2F2E9 | Saved revisions, clean recovery state |
| Secondary/disabled | Blue-gray #E5ECF3 | Secondary tools; disabled actions use muted text |

Color reinforces existing labels, button availability and bold state; it never changes authority or report data. Native table rows retain their status words rather than implying per-row coloring unsupported by the current list control. Normal text palette pairs measure at least 5.92:1 contrast; disabled controls are muted. These measurements do not establish full accessibility conformance.

## Visual findings and adjustments

- Checked the Overview, Reports, Summary groups, My Work, Resources, Settings and pop-out guide on synthetic data in the visible full Access application.
- Corrected active-navigation emphasis; Overview no longer stays emphasized in another workspace.
- Corrected the rail stacking order so its opaque background stays behind navigation controls.
- Enabled native custom-color command-button rendering and assigned matching hover/pressed text/background colors.
- Widened report status/promotion columns and group headings. Wide lists may use horizontal scrolling rather than truncating key status headings.
- Removed the Resources menu gap left by hidden legacy guidance.
- Replaced comma-containing help topic names with concise complete titles so Access's value list does not truncate them.
- Inspected the fixed report editor. Moved the reusable-senior picker out of the page-tab/header area into the footer, preserving the function. Shortened calculation summaries while keeping full values in tooltips. Official report field positions and seven Front Page traits remain unchanged.
- Live save/recovery, validation and queued-cell feedback uses the same palette; field-level validation colors retain their own renderer.

## Boundaries

Visual observations cover the current display/window configuration. Native compact-layout checks supplement the screenshots; physical 125/150 percent scaling, keyboard/screen-reader acceptance and exact printed fit are not claimed. No Windows display settings were changed. Full release migration and standard form-generation acceptance are separate gates.

Candidate and final test/hash evidence are recorded in VISUAL_REVIEW_VERIFICATION.json. Original populated workspaces and earlier candidates are preserved. Intermediate visual previews are synthetic scratch artifacts, not release candidates.

Final exact-artifact results: 107 palette/navigation/status/geometry assertions, 57 help assertions and 264 fixed-page assertions passed. Structured smoke Status=PASS matches the final artifact hash. Ten installed modules match source. Final candidate: build/visual-checked.accdb; blank manifest: build/visual-checked.manifest.json. Earlier intermediate visual builds also passed 30 navigation, 24 header-review and 42 pending-cell regressions; those are supplementary results, not substituted for the final exact-artifact checks.

All 68 form definitions accept the palette without property errors on a disposable copy. This is a property-application check, not a claim that every form was visually inspected.
