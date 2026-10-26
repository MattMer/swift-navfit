# SV-7 Systems Measures

| Measure | Threshold | Current | Evidence |
| --- | --- | --- | --- |
| Historical import coverage | Every prior schema imports with every row and field verified | Schemas 1Ã¢â‚¬â€œ34 | `test-swift-data-import` |
| Regression suite | Every core and extended runner passes before a release | See the counts below | `develop.ps1 -Action ExtendedTest` receipt |
| Text contrast | 4.5:1 or better for every palette pair | 12 of 12 pairs | `test-accessibility` |
| Accessible names | Every input control and button named | No known gaps | `test-accessibility`, `tests/accessibility-gaps.json` |
| Very wide windows | Layout completes at Access's form-width limit | 36,000-twip check passes | `test-visual` |
| Resize cost (M09) | Same-size resize does no layout writes | My Work 68 Ã¢â€ â€™ 12 ms; resizing to a new size about half its earlier cost | M09 benchmark (commit a943e60) |
| Printed fit | Every block fits the official printed form | Not yet measured | Backlog H01 |
| Screen reader | Every step of the attended procedure passes | Not yet measured | [Screen-reader test](../../Accessibility_Screen_Reader_Test.md) |

<!-- generated:counts -->
Test runners by tier (`tests/runners.json`):

| Tier | Runners |
| --- | --- |
|  | 1 |
<!-- end generated:counts -->
