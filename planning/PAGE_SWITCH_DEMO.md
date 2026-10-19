# Front/Back switching and validation demo groups

2026-10-19. Full desktop Access development candidates.

## Changes and measured result

The submitted 9.2-second recording shows populated FITREP page transitions. The old code moved/hid controls while Access could repaint, refreshed the surface twice, repeatedly assigned unchanged geometry/colors, and reset review colors before applying them again.

Page layout now suspends parent/surface painting as one batch and restores both previous states on success or failure. Active text is flushed before focus moves. Unchanged geometry/palette properties are left alone, duplicate refresh is removed, and review colors have one owner. Saved values, locks, revisions and recovery remain authoritative; official report fields and all seven Front Page traits retain their geometry.

Equivalent populated synthetic FITREP copies, 10 alternating switches each: median 771 -> 373 ms (51.6 percent reduction). Back medians 598 -> 281 ms; Front medians 805 -> 422 ms. These are local measurements, not a universal guarantee. The optimized populated FITREP was visibly switched Front -> Back -> Front in full Access. Paint restoration is also tested after an injected layout failure; no high-speed flicker metric is claimed.

## New demo groups

- DEMO - Clean reports (0 report issues): Morgan Ellis and Taylor Hayes; zero member-report findings.
- DEMO - Practice corrections (3 report issues): Cameron Avery has missing command employment, duty description and PFA/readiness summary. Reese Parker is issue-free.

As requested, standard roster-completeness and provisional-capacity reminders remain: full group lists contain 2 findings for Clean and 5 for Practice. Validation rules were not weakened. The seed asserts clean baselines and exactly three deliberate report issues and refuses populated destinations.

The demo now has 13 synthetic reports, 5 groups and 12 prior evaluations; original review/approval/archive/NOB/recovery examples remain. Existing demo files were not overwritten.

## Verification and handoff

29 active-text/page-switch/paint restoration assertions, 264 fixed-page assertions, 107 palette assertions, 12 recovery preparation and 23 recovery assertions plus timer/reopen checks, and 37 populated-demo assertions passed. Application smoke passed against the exact blank-candidate hash. All three installed modules match source. Blank and populated demo inventories are separate.

The demo harness now tests 13 reports/5 groups and the live member window instead of obsolete hidden-workbench assumptions. The standard builder includes the update; normal demo generation uses the expanded seed.

- Demo: build/page-switch-demo.accdb
- Blank candidate: build/page-switch-candidate.accdb
- Evidence: PAGE_SWITCH_DEMO_VERIFICATION.json and build/page-switch-performance.json

The populated r116 original retains its recorded hash. The earlier visual-checked development candidate differs from its old manifest hash; cause is undetermined and it is not labeled unchanged. Copy-based builders verified their current template hash before/after, and the new artifacts have independent exact-hash inventories. No populated operational workspace was replaced.

Physical display scaling, keyboard/screen-reader, print-fit and full release-migration acceptance are not claimed. Free Access Runtime support is not required.
