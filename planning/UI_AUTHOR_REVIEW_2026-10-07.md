# Author-edited baseline — r89 (7 October 2026)

Use `output/ux/Swift_NAVFIT_UI_Demo_2026-10-07-r89.accdb` for continued UI work. Matching blank: `output/ux/Swift_NAVFIT_UI_Blank_2026-10-07-r89.accdb`. Version 0.25.0-dev/schema 24; published GitHub release remains unchanged. The supplied `Swift_NAVFIT_UI_Demo_2026-10-06-EDITED.accdb` and r88 are preserved.

## Reviewed changes

Seven forms changed: frmLibrary, frmMajorWorkbench, frmMyWork, frmUXGroups, frmUXImportChoose, frmUXMemberHistory and frmUXResources. There were 316 property/object differences, with no standard VBA module or form-class code changes. Adopted tighter report filters/actions, revised Validation/Move report/Summary Group wording, the new Column list and Archive status labels, grouped summary-group actions/defaults, the My Work contributor placement, the shorter history explanation and stronger member heading, and import/resources alignment changes.

The baseline is reproducible through `src/ui-author-baseline.json` (303 property changes and two added labels), applied by `scripts/apply-ui-author-baseline.ps1` at the end of the current full and incremental build path. Runtime My Work layout preserves the author's contributor placement at normal widths, while retaining the compact-width layout. Raw report-list ColumnWidths are not adopted because the nine-column adapter computes them from the selected view and must keep internal IDs hidden. Some conditional workbench controls intentionally retain their runtime positions when their tool opens rather than displaying the author's off-screen design positions.

## Hidden-control cleanup

The r88 inventory contained 524 hidden controls. Traced source references, event bindings, tags, data sources and dynamic families before selecting removals. Removed 50 additional obsolete controls from the author's copy; including eight authored deletions, the candidate has 58 fewer obsolete controls than r88. These include old Home cards/actions, duplicate editor labels/detail buttons and old bulk page buttons. The exact allowlist and reasons are in the baseline JSON.

Retained unbound report/bulk buffers, IDs/stamps, conditional section/error controls, responsive navigation, trait labels addressed by constructed names, reset/inheritance buttons and reporting-senior adapter controls. Restored only `frmLibrary.cboDestination` as a hidden compatibility buffer: older folder functions still reference it. No tables, forms, modules or backend services were removed.

## Verification

Compiled both candidates. All 49 user tables retain identical field schemas, row counts and row-value hashes compared with the author's returned database, including its additional report, group, revisions and note. Both originals' hashes stayed unchanged during inspection.

Passed 2,684 assertions: 142 admin/comparison, 2,107 responsive/focus/identity, 22 group organization, 32 My Work, 355 author-property/cleanup/startup/populated-navigation and 26 approved-trait retry checks. Blank normal startup displays Setup; the demo startup check displays Overview. Native visual spot-checks covered Reports, Summary groups and My Work. This does not establish screen-reader or physical DPI acceptance.

The supplied file carries Zone.Identifier. It was preserved on the original; removed only from the newly generated, reviewed local build. Access startup protection remains intact. No global trust or security setting changed.

Two visual follow-ups remain: report columns truncate some headings/names at the observed window width; the Summary groups archive-state filter says Group status even though the workflow status selector is separate below. Recorded for the next UX pass rather than changing the author's design beyond the requested cleanup.

## RUAD

Received and inspected the synthetic sample. It has 35 columns, 66 rows, 60 named members and six vacancies, with no malformed rows. Mapping details and identity/refresh behavior are documented in `RUAD_MAPPING_2026-10-07.md`; counts in `RUAD_SAMPLE_ANALYSIS_2026-10-07.json`. Earlier sample-blocked notes are historical. RUAD import and refresh implementation is the next remaining feature; it is not included in r89.

Evidence: UI_AUTHOR_REVIEW_VERIFICATION_2026-10-07.json and UI_AUTHOR_DATA_PRESERVATION_2026-10-07.json. Use the current full builder or build-ux-admin-feedback.ps1; historical incremental recipes do not represent the new baseline.
