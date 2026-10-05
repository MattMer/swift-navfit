# Swift NAVFIT redesign — prototype review

The distributed Access application remains version 0.21.0, schema 20. Native redesign development is version 0.22.0-ux-preview, schema 21. The clickable design prototype remains a separate artifact containing synthetic data. It does not connect to Access, read source files, certify Navy policy, or implement a replacement web application. Native user-review builds now include their own synthetic examples through `scripts/build-preview.ps1`; they do not share browser storage.

## Open the prototype

Run `scripts/build-ux-prototype.ps1` to generate `output/ux/Swift_NAVFIT_UI_Prototype.html`. Open that self-contained HTML file in a browser; no server is needed. The development version in `prototype/` can also be served with a local static HTTP server. Browser demo changes persist in local browser storage only. Settings & Help can reset the examples or create an empty demo workspace.

## Review journeys

1. Reports → open Taylor, Jamie. Move through official report sections, type into fields, then switch sections. Pending values should remain. Save creates a synthetic revision; the separate recovery timer does not.
2. Enter `2026-02-30` as a period date and an overlong official name. Save should show a clickable error summary while preserving all entries. Correct them, save, and return to the filtered report list.
3. Open Lee, Jordan and its Review section. Edit a report field, then start a review. Inspect the choice between saving pending values and reviewing the saved revision.
4. My Work → Reed, Casey. Review the stale-route explanation. Open Patel, Sam to inspect local approval locks and reasoned reopening.
5. Reports → select editable reports → Bulk change. Inspect full before/after values, apply, and preview guarded undo. Selecting an approved report should explain why the batch cannot proceed.
6. Summary groups → Members, Shared defaults, Review, Comparison and History. Open Comparison directly, change its presentation order and preview trait edits. Saved-value baselines are captured automatically. Default changes should not alter report values automatically.
7. Reports → Import. Compare the narrative-only NAVFIT scope with the eight-field interactive-PDF preview and required interpretation choices.
8. Overview → recover pending edits. Check that recovered text is still unsaved. Settings & Help → empty workspace tests the starting experience.

## Design decisions implemented

Balanced Reports / Summary groups / My Work navigation; official block order with form-specific narrative blocks; explicit save plus recovery; synthetic examples for all three form kinds; named local contributors and contextual actions; laptop-oriented collapsible navigation; visible focus and keyboard-accessible native browser controls; technical detail moved to Settings & Help.

All 78 existing draft fields are represented in the prototype metadata. Separately stored names do not overwrite one another. Labels and block associations still require a final current-template review during native implementation. Existing trait catalogue captions/guidance are reused for the prototype, with their current catalogue provenance rather than new policy claims.

## Deliberate prototype boundaries

Revision comparison/restoration, route-template editing, complete provenance, closed-file backup and the Navy rule engine have destination previews rather than replicated backend implementations. The native services already implement these capabilities and must be retained. Browser storage is not a migration or recovery guarantee for the Access workspace.

Browser inspection through the available automation tool was blocked because saved browser permissions could not be verified. Syntax/model checks do not establish visual, keyboard, screen-reader or DPI acceptance. Review in a browser is still required before native rebuilding. Native Access keyboard/screen-reader, display/scaling and migration acceptance remain future verification steps.

## Next implementation stage

Resolve prototype navigation, wording and density feedback first. Then rebuild shared native chrome and the section-based unbound editor while retaining the existing field control identities, save/recovery baseline and DAO services. Replace modal detail entry points with section navigation, split the workbench into contextual destinations and add native dirty-state/inline-error handling. Keep schema 20 and ship a separately verified blank prerelease.

## Review feedback incorporated

Official block captions and form-specific block numbers now drive the prototype. The normal editor uses SSN; retained DoDID values remain outside the normal UI. Occasion, NOB and report-type flags use checkboxes. The editor displays pending ITA and the selected group average from saved values. Summary groups open from a group list into a focused workspace. Shared group changes use report sections, and comparison shows compact editable trait columns without narratives.

The revised prototype passes 102 model/controller/calculation checks plus actual ES-module loading/navigation regression checks. Native Access construction is underway, with verification still pending; it is not yet a release candidate. Browser automation was unavailable, so automated logic checks do not establish visual or accessibility acceptance.

## Reporting seniors and PSG RSCA

Resources now contains a reporting senior directory with all 16 reusable saved senior fields, including SSN, structured name parts and address/contact details. Full names and name parts remain independent. Reuse is explicit: pending values in one report, or a full before/after group preview.

Rank-specific immutable historical baselines retain RSCA, previous graded trait count, optional exact trait sum, as-of date and source note. Each group explicitly selects its reporting senior and a same-paygrade baseline and confirms that the totals exclude the current group. PSG RSCA uses (historical trait sum + current group numeric trait sum) / (historical trait count + current group numeric trait count). Blank/NOB traits and whole-report NOB do not contribute; blank/partial group context is disclosed. Rounded RSCA times count is visibly labelled an estimate when exact historical sums are missing. Calculations do not advance the baseline or insert comments into reports.

Reference: MyNavy HR, Administrative Officer Guide - Managing RSCA (3 June 2024), https://www.mynavyhr.navy.mil/Portals/55/Career/Detailing/Officer/SurfaceWarfare/Documents/1610-1%20AO%20Guide%20-%20Managing%20RSCA.pdf . This guide calls for the historical Total Trait Sum and Total Trait Count from BOL, along with the current group trait marks.

Schema 21 is a justified data requirement for the native directory: four additive tables retain reporting senior records/values, historical RSCA baselines and explicit group-baseline selections. The 78 report fields remain unchanged. Existing workspaces are transferred into a new blank target; source files are not modified. Native visual, accessibility and DPI verification remains pending.

## Native verification checkpoint

0.22.0-ux-preview, schema 21, `build/ux-candidate9.accdb`, SHA256 `34C96AD947DCC933F995053CF572E165544CB750412D889240F73ED3B5480F0B`. All 42 user tables are blank. The 24 existing regression suites passed candidate8; final candidate9 passed 38 reporting-senior/RSCA/comparison checks, six workbench UI checks, 65 visible control-bound checks and 34 recovery assertions. A guidance-lock regression exposed by the broader suites was repaired.

Schema 20 to 21 transfer retained a report, all 78 revision values and two recovery checkpoints. Populated schema 21 to 21 transfer retained five revisions, 390 revision values, a 16-field reviewed batch, the senior record, four historical RSCA baselines and group associations. Row/field comparisons, unchanged source SHA256 and reopen checks passed.

This is a development checkpoint, not the published blank prerelease. Native screen layout, keyboard/screen-reader behavior and scaling still require manual verification, and the remaining native redesign acceptance work is ongoing. The distributed 0.21.0 binary and its release manifest remain unchanged.
