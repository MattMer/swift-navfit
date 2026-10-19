# Recent Swift NAVFIT development

Source update: 19 October 2026. Development version 0.25.0-dev, schema 25. These source changes do not constitute a new binary release. Existing downloads and historical verification notes describe their own exact artifacts.

| Area | Change and verification notes |
| --- | --- |
| Reports and My Work performance | [Batch 1](PRIORITY_CLEANUP_BATCH1.md), [display reads](PRIORITY_CLEANUP_BATCH2.md), [group calculations](PRIORITY_CLEANUP_BATCH3.md), [member-grid reads](PRIORITY_CLEANUP_BATCH4.md), [queue inclusion/counting](PRIORITY_CLEANUP_BATCH5.md) |
| Save and recovery feedback | [Batch 6: saved revisions, pending work, scope and contextual actions](PRIORITY_CLEANUP_BATCH6.md) |
| Navigation and help | [Canonical contextual guide](NAVIGATION_HELP_CONSOLIDATION.md), [explicit defaults context and comparison views](NAVIGATION_UPDATES.md) |
| Visual presentation | [Semantic palette, active navigation, menu spacing and editor header/footer](VISUAL_REVIEW.md) |
| Front/Back switching | [Batched redraw and measured performance improvement](PAGE_SWITCH_DEMO.md) |
| Demo scenarios | [Clean-report and three-issue groups](PAGE_SWITCH_DEMO.md); standard group reminders remain |
| Wider report view | [Available-width viewport and narrow-window scrolling](PAGE_WIDTH.md); fixed field positions retained |
| Build fixes | [Native unload guards, property binding and finalization](BUILD_SMOKE_FIX_2026-10-19.md) |

Each note names its exact development artifact, checks, and limits. `build/` artifacts and scratch files are local and excluded from Git. The populated r116 workspace must be preserved; a filename containing “Blank” does not certify empty tables. Use the read-only artifact inventory before distributing any purported blank build.

The target is full desktop Microsoft Access. Free Runtime support is not required. Release-migration acceptance, physical display/scaling, keyboard/screen-reader coverage, and final printed fit must be assessed separately from targeted native tests.

Publication checks: the latest widened blank candidate passed the repository lifecycle/review workflow suite (40 native-close, direct shared-field and report-review assertions, including the script's forced-quit checks). Changed PowerShell and JSON files parse, README local links resolve, and the staged diff passes whitespace checks. Binary databases, recordings and local scratch helpers are excluded from this source update.
