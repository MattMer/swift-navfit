# Native layout checkpoint - 2026-10-04

## Latest checkpoint: denser sections and visible group counts

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r13.accdb`.
SHA256: `937C79C3D4CA64FC119694F923D7CC3B29C338E7E5281D8B05777F1ADEB86FC3`.
Older preview files and their edits remain unchanged.

All official sections use closer label/input spacing and shorter row pitches.
Checkbox captions get sufficient horizontal room, and empty inline errors are
hidden. Actual errors expand their row so messages do not overlap later fields;
showing a save-error summary also moves the field panel below the summary.
Trait captions and compact mark selectors share rows, with fixed adjacent
guidance. Unused vertical extents are reset to allow all seven traits to fit
without blank scrolling space from hidden narrative controls. Review actions
share one compact row; History actions sit above a viewport-sized revision list.

The report-count query already existed, but total column widths exceeded the
Summary groups list width and hid its last column. The visible final column is
now named Reports and counts current GroupMember records (including NOB and
archived reports that remain members). It is not a saved comparison roster count.
Widths now fit the list. Existing group-load/activation refresh behavior remains.

Generated label punctuation is normalized to ASCII before Access imports the
VBA source, correcting the UTF-8/ANSI corrupted readiness captions seen in the
user screenshots. Labels size to their captions to reduce clipping.

Verification: 108 surface/workspace assertions, including tight field gaps,
compact seven-trait layout, Review/History positioning, captions, group counts,
validation and recovery; 33 populated demo assertions; both startup paths.
Native visual/DPI and screen-reader acceptance remain pending.

## Previous checkpoint: compact report block rows

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r11.accdb`.
SHA256: `2BCBE20517D1B2ABD44383167697F0C2E55931113AB1EE48210934A033AE7961`.
Previous preview files remain unchanged.

Equal-size field cells are replaced by content-sized rows. Component flags,
occasion flags, NOB/report-type flags, retention choices and statement choices
share compact rows with checkbox captions beside their controls. Member details
follow report block groups: name/rate/designator/SSN, component, command details,
occasion, period, report kind and readiness. Structured name parts remain a
separate compact row within the Member section; stored full names remain independent.
Short fields use narrower widths, long names use wider widths, narratives span
the panel, and trait guidance retains its adjacent area. Rows wrap when needed.

The supplied current CHIEFEVAL PDF is dynamic XFA, so the available page renderer
shows its placeholder. Read-only XFA template inspection confirmed grouping and
relative widths (including 120.65 mm name versus 19.05 mm rating, clustered
component flags, occasion/period and NOB/report-type areas). This is interaction
layout inspired by report blocks, not a claim of verified printed layout.

Verification: 101 native surface/workspace assertions, including compact checkbox
rows, adjacent labels, resizing, input bounds and preservation of pending text;
32 populated demo assertions; normal-security and enabled-content startup checks.
Manual native visual, DPI and screen-reader acceptance remain pending.

## Previous checkpoint: report editor uses available window space

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r10.accdb`; earlier user files
are preserved. SHA256:
`A1AD37356B3836446FB8B08E279AFFC3030E51D63C2814DBE3FA9ABBC4CC7DA2`.

The editor handles its Resize event with explicit viewport geometry. The field
panel fills the content area's width and remaining height, with fixed-height
save actions near the bottom. Field controls reflow to two columns on narrower
windows and three on wider windows. Narratives span the panel; trait guidance
uses the adjacent column. Section buttons and the summary-group selector have
more room, and the approval warning spans the header instead of a narrow rail.

Hidden control extents (including a 12,800-twip heading) previously constrained
the child form width. Reflow now resets unused horizontal extents before sizing
visible controls and the child form. Resize flushes active text into the unbound
buffer, preserves all pending values and creates no revision.

Verification includes actual Access maximize/restore, narrow/wide viewport
fixtures, pending text preservation during resize, horizontal section bounds,
32 populated demo journeys, and both startup security paths. Full physical DPI,
keyboard and screen-reader acceptance remain pending. Other destination forms
still need their responsive layout pass.

## Previous checkpoint: startup guard and stable editor placement

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r7.accdb`. This is a clean
source build with the same nine-report synthetic dataset; previous previews
and their user edits are preserved. SHA256:
`04C79F733CF62F3F02A77762ED52032951D34361154E7CA475527B210BFA3A1F`.

The reported launch failure was reproduced with normal Office security:
active content was blocked. AutoExec now checks CurrentProject.IsTrusted;
blocked content opens a code-free instruction form rather than invoking VBA.
Enabled content calls a distinct SwiftAppStartup entry point, with useful
startup error reporting. A scoped registry trust setting was attempted but
was ineffective in the automation test and removed; Office settings remain
unchanged. Enable Content and reopen, or add the preview folder in Trust Center.

The field panel's right/bottom anchors and the footer's vertical stretch were
incorrect. Left/top anchors retain the fixed laptop layout and 500-twip button
heights across maximize/restore. The contributor header now updates loaded
forms directly, and section/navigation captions escape Access ampersands.
This fixes placement; responsive collapse and DPI acceptance remain pending.

Verification on the clean build: 80 surface/workspace assertions, 32 populated
journey assertions, normal-security AutoExec and repeat launch, enabled-content
startup into Overview, and blank-template startup into Setup. Runtime checks
cover maximized/restored field placement, footer size and non-overlap. This is
not manual visual or screen-reader acceptance. The preview builder now requires
startup verification as well as pre/post-seeding behavior checks.

## Previous checkpoint: native workspace rebuild with synthetic demo

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-04-r4.accdb` for the previous user
preview. It contains nine reports, three groups, two contributor profiles, one
reporting senior with three RSCA baselines, and review/approval/archive/NOB/
recovery examples. First-run setup is already complete. Demo SHA256:
`06BF3D9EC4E6DF008159F6C301998D169BD28008E377A1E28B546E3BFE95F124`.

The blank verified source is `build/ux-candidate21.accdb`. Core changes:
- Complete official editor sections replace six-field pagination. An unbound
  child surface scrolls independently while the parent header, navigation and
  save area remain present. The original 78-field buffer/services are retained.
- Visible fields bridge to the buffer; active text is included in checkpoints
  and saves. Section navigation preserves pending values without revisions.
- Field labels follow the form type, including comments blocks 43/40/41.
  Checkbox selections map to the existing Yes/No/null and NOB storage values.
- Inline input errors and a selectable summary retain pending values and open
  the relevant section. Incomplete drafts remain supported.
- Overview uses real metrics and recent reports. Reports gains type/status
  filters, official-name search and restored selection after refresh.
- Group lists open a focused workspace with member, average and PSG context.
  Comparison separates trait editing and staged-proposal preview.
- My Work separates assigned work, assigning a saved report, and contributors.
- `build-preview.ps1` creates seeded previews and runs pre/post-seed checks;
  `build.ps1` remains the separate blank-template builder.

Verified on candidate21 and the finished demo's scratch copies: all 24 existing
regression suites, 73 new native surface/workspace assertions, 38 reporting
senior/RSCA/comparison assertions, 32 populated-demo assertions, six workbench
UI assertions, 425 main-form bounds checks and 76 additional bounds checks.
Source hashes remain unchanged by preview generation and scratch testing.

The first section surface exposed Access's focused-control visibility rule;
switching now moves to a permanent focus anchor before hiding child fields.
The VBA variable `tab` was also renamed to avoid a reserved identifier. These
failures reinforced the need for execution checks rather than compile-command
success alone. Old detail-dialog layout tests now check overlap only when their
legacy entry button is actually visible; their data/recovery tests remain.

Visual parity, DPI/collapse, native keyboard/screen-reader acceptance, remaining
review/import/bulk screen refinement and native close-window dirty guards are
still pending. This is an intermediate preview, not a published replacement.
See [demo guide](UX_DEMO_BUILD_GUIDE.md).

## Previous checkpoint: candidate13

The user correctly identified that candidate9 did not resemble the reviewed
browser prototype. It was a partial functional adapter, not a faithful native UI.
Do not describe it as the implemented redesign.

Candidate13 adds a consistent left navigation rail to thirteen main workspaces,
selected destination styling, Segoe UI, content bounds, an Overview with four
journey cards, and a Reports folder filter instead of the legacy folder tree.
The report type selector is visible. Existing service-facing control identities
and editor buffer remain intact. Existing workspaces are not modified.

The new folder-filter callback was initially inserted ahead of module-level
declarations, breaking their scope. The corrected module places all declarations
before procedures. Access RunCommand(126) alone failed to establish executable
validity; behavioral execution must accompany build checks.

Verified on build/ux-candidate13.accdb:
- 38 reporting senior, RSCA and comparison assertions.
- 34 recovery assertions.
- 6 workbench UI assertions and 75 visible-control bounds checks.
- 46 smoke assertions, including setup, three report types, atomic saves,
  archive guards, folder filtering, editor actions and reopening.
- 560 navigation/content bounds checks.

This remains an intermediate development artifact. It has not been visually
verified against the prototype, checked at Windows scaling levels, or tested
with a screen reader. Native editor pagination, the Overview's missing recent
reports, long workflow screens, inline validation, and persistent navigation
while vertically scrolling still require redesign. Bounds checks do not prove
readable labels, absence of overlap, or visual fidelity. The rail is currently
within form detail, rather than a fixed shell hosting destination content.

The published 0.21.0 release and existing workspaces are unchanged.
