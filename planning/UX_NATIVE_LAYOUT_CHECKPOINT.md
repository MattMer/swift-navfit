# Native layout checkpoint - 2026-10-04

## Latest checkpoint: startup guard and stable editor placement

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
