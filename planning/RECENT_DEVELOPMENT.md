# Recent Swift NAVFIT development

Current delivery: r128 / 0.27.1-dev, schema 27. See [PROJECT_STATE.md](../PROJECT_STATE.md) for current architecture and [release history](CURRENT_STATE.md) for exact-artifact evidence. Older entries below describe their own checkpoint.

| Area | Change and verification notes |
| --- | --- |
| Startup contributor selection | [r128 explicit session choice, attribution and verification scope](CONTRIBUTOR_STARTUP_R128_2026-10-21.md) |
| Reviewed duplicate Sailor merging | [r127 preview, identity guards, retained history and undo](DIRECTORY_MERGE_R127_2026-10-21.md) |
| Local PDF input | [r126 local extraction and supported-form limits](LOCAL_PDF_R126_2026-10-20.md) |
| Development consolidation | [Current workstream decisions and verification](CONSOLIDATION_PLAN.md) |
| Build and legacy cleanup | [Managed build pipeline and retained adapters](BUILD_PIPELINE.md); [r125 implementation and verification scope](BUILD_CLEANUP_R125_2026-10-20.md) |
| Workspace import verification | [r124 database-engine exact verification](ENGINE_IMPORT_R124_2026-10-20.md) |
| Consolidated report tools | [r123 notes, replies, review, validation and calculations](REPORT_TOOLS_R123_2026-10-20.md) |
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

Historical widened-candidate publication checks included 40 lifecycle/review assertions. That checkpoint is not current release acceptance; use the exact release receipts above. Binary databases, recordings and local scratch artifacts remain outside maintained source.
