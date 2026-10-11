# Standardized narrative editor — r104 (10 October 2026)

## Current candidates

- Synthetic demo: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-10-r104.accdb`. The first EVAL includes the supplied Lorem Ipsum sample (Jamie Taylor in the existing synthetic fixture).
- Verified blank template: `output/ux/Swift_NAVFIT_UI_Blank_2026-10-10-r104-clean.accdb`.
- `Swift_NAVFIT_UI_Blank_2026-10-10-r104.accdb` retains the data copied from the user's locally populated r103 file; it is not the blank distribution.

App 0.25.0-dev/schema 25. Existing files and the GitHub r103 asset are preserved. Published as the r104 development prerelease with blank and synthetic-demo assets.

## Reference and implementation

The supplied EVAL page-2 PDF contains Courier New 9-point narrative text, 18 populated lines, about 10.2-point baseline spacing, and up to 92 characters on a displayed line. This confirms that NAVFIT's displayed “10 POINT” option is not the PDF's actual 10-point font. These observations are a sample, not a universal line or character limit.

The comments editor uses plain text, locked Courier New, normal weight, fixed margins and a fixed report-type footprint. EVAL/FITREP use 9-point text with a 7.15 by 2.60-inch text rectangle; CHIEFEVAL uses 10-point text with a 7.60 by 1.15-inch rectangle. Controls add 0.0625-inch margins on each edge. Rectangle references come from the supplied NAVFIT98A 2.2.0.332 default print code; only EVAL has equivalent supplied output corroboration. See NARRATIVE_LAYOUT_R104_MEASUREMENTS.json and NAVFIT_PRINT_ENGINE_FINDINGS.md.

General field-grid resizing no longer stretches comments. Navigation collapses when needed to make room. Very narrow viewports shrink the field without changing its font and explicitly display “Narrow view: wrapping differs.” The character count refreshes with the existing change-driven validation timer and includes active typed text.

Overflow remains scrollable and fully saved; no truncation or guessed character cap is introduced. Layout/section changes create no revisions. Other multiline blocks keep their existing formatting.

This standardizes the input but does not certify printed fit. Access and NAVFIT use different rendering/wrapping engines, and NAVFIT's PDF shows horizontal text scaling. No saved CommentsPitch/import provenance or export behavior was changed. A returned/imported report's retained pitch may differ from this default editor reference. Equivalent CHIEFEVAL/FITREP samples, alternate pitch paths, printer comparisons and explicit measured overflow/preview remain future work.

## Verification

Compiled both native candidates. Native scratch tests passed 129 assertions across all three report types: fixed geometry/font/margins, three normal viewport widths, narrow-view disclosure, no overlap with inline errors, section preservation, unchanged revisions on layout, and 6000-character overflow save/reopen. Automated properties and behavioral checks do not establish physical display or exact wrapping acceptance. Existing narrative regression checks and blank startup are recorded in the verification JSON.
