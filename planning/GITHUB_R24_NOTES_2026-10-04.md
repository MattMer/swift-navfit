# GitHub source milestone: r24 — 4 October 2026

This cumulative source update brings the outstanding feature work and native UI redesign onto the main development line. Swift NAVFIT remains version 0.22.0-ux-preview, schema 21. This is a development preview; merging source does not publish a new Access binary release or replace an existing workspace.

## Current behavior

- Persistent Overview, Reports, Summary groups, My Work, Resources and Settings & Help navigation. Reports use official block sections, compact layouts, SSN, checkbox selections and visible trait averages. All 78 saved report fields remain available.
- Reporting-senior directory, manually entered rank-specific RSCA baselines and group PSG calculation context. Group counts, lifecycle actions, shared-block bulk previews and editable trait comparison are integrated into their workspaces.
- Automatic save on normal departure, explicit Save draft, independent recovery checkpoints, failure-preserved pending edits, atomic revisions and stale-baseline guards.
- Guided import, native workflow help, blank/demo packaging and source-preserving workspace transfer tools. Existing provenance, guarded undo, approval and recovery services remain in use.
- r24 fixes the top-of-list report reopening crash. My Work combines assigned reviews, current ordered-route steps, ready local approvals and related unresolved threads, with named reports, saved revisions, next actions and disabled stale/locked decisions. Routes/comments have no invented dates or assignees.
- Shared title/filter spacing, visible Save captions, full-width lock guidance, readable list headings and supporting-screen navigation are improved. Unknown folder text is rejected without changing the current filter.

## Verification

Both final local artifacts compile and match all 50 source modules. The blank has 42 empty user-data tables. The demo has nine synthetic reports, three groups and six initial inbox items.

Passed checks include 145 focused report-return/filter assertions, 32 unified inbox, 38 routing and 28 comment assertions. The preview gates also passed imports, help, calculation context, group lifecycle, editable comparison, responsive layouts, editor buffers, save/bulk/review behavior, populated demo and startup checks. Forced Access shutdown tests verified valid saves and preservation of invalid raw edits for recovery. PDF integration and the full foundation suite were not all rerun for r24.

Disposable native checks covered single-result opening, current/stale work selection, opening the correct thread, advancing a route to the next contributor and desktop/laptop layouts. Full evidence and final artifact hashes are in [UX_MY_WORK_2026-10-04.md](UX_MY_WORK_2026-10-04.md).

## Outstanding and next focus

1. Complete native mutation journeys, keyboard/screen-reader acceptance and actual Windows scaling at 100%, 125% and 150%; verify Runtime/x86 and exact release packaging.
2. Polish supporting-screen density, accumulated Access tabs, object-pane exposure, legacy indeterminate checkbox values and trait guidance.
3. Finish controlled NAVFIT interoperability mappings, export/returned-file conflict handling and policy-defined grouping/validation proof. Printing remains deferred.

No network integration, authentication, synchronization or official signatures are added. Local approval remains a local workflow lock. See [CURRENT_STATE.md](CURRENT_STATE.md) and [PROJECT_PLAN_STATUS_2026-10-04.md](PROJECT_PLAN_STATUS_2026-10-04.md) for the complete plan.

Generated Access files, installed NAVFIT files, supplied PDFs, machine-specific evidence and credentials remain excluded from Git. Use `scripts/build-preview.ps1` for synthetic demos and `scripts/build.ps1` for blank builds. Existing release assets remain the previous distribution checkpoint.
