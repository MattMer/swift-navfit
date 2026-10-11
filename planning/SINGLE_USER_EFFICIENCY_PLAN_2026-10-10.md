# Single-user efficiency cleanup - 2026-10-10

Source review against r100. Identification and plan only; no runtime changes in this review. Keep one user per workspace, with exclusive opening enforced before removing assumptions about external writers. Current generated setup says exclusive use is not enforced (scripts/build.ps1).

## Recommended order

1. Separate live Members/trait editing from immutable meeting/proposal baselines. modUXTraitGrid.UXGridEdit currently calls FreezeSummaryGroup and CreateComparisonSet after every successful cell save, copying every member twice. Ordinary Members opening should read current saved reports directly. Create history baselines only for explicit proposal/comparison operations that require them. Preserve existing historic proposals, report revisions, atomic saves, queued retry/discard and approval reopening. Target: no group-history rows for opening Members or changing a direct-edit cell.
2. Replace display-time database guards with action-driven state. UXGridLockReason is called for every loaded row, on row focus and again at save. It repeatedly reads the report, snapshot and set and computes the whole-group GroupStamp. Load report archive/workflow/revision state once with member rows; update after successful local mutations. Do one authoritative check at save/apply, not at every focus change. Keep historical proposal, import merge, restore and saved-revision approval checks.
3. Target editor refreshes. UXRefreshEditor requeries the group selector and reads group averages after each field change. Refresh membership/default selectors on open or explicit group mutation. Recalculate pending ITA only after trait/observation changes and saved SGA after saved group-report changes. UXSurfaceEdited and UXFieldChanged refresh chains need deduplication. Retain immediate field validation and pending-value preservation.
4. Remove redundant legacy comparison work only after dependency proof. UXLoadComparison fills lstTraits (seven per-member revision lookups) alongside the editable grid, and populates/requeries lstProposals before UXRefreshComparisonPreview repopulates it. Confirm generated events and proposal tools that still use lstTraits/old controls; adapt them before removing controls or queries. Collapse duplicate proposal population now; lazily load proposals when their tool is requested.
5. Simplify group page navigation. UXGroupNavigate coordinates frmMajorWorkbench as a hidden context host plus separate comparison/default forms, repeats member counts, and closes/reopens forms when changing pages. Introduce a group-context owner, fetch the count once, and route directly to the selected group page. Preserve pending-cell and bulk-preview exit handling. Members/Comparison already share a form; avoid duplicating that work.
6. Idle editor timer: after mutation-driven invalidation is established, avoid scanning all 79 fields every second. Field OnChange marks validation dirty; OnAfterUpdate/save flushes validation; timer handles typing debounce and changed recovery checkpoints only. Keep active-text capture and failed-write retry.

## Checks that still have a purpose with one user

- Save/recovery baseline checks: the same user can change saved data through another form or an import while an editor has pending data.
- Proposal/import/bulk preview baseline checks: previewed values must still match the intended saved revision when applied.
- Approval saved-revision linkage, workflow/archive restrictions, duplicate member-period checks, input/policy validation, transactions and rollback.
- Pending save failure guards: these prevent dropping unsaved queue values; they are not concurrency guards.
- Recovery session/contributor attribution and checkpoint existence checks.

Do not remove a check purely because its wording says changed or stale. Remove repeated read-only evaluation from display paths, retain a bounded check at the relevant write boundary.

## Verification targets

Count history rows and SQL/refresh calls across group opens, row focus, single-cell saves, editor field changes, idle time and group tab switches. Test every report type, approval reopen/rollback, pending-cell failures, import while another form is open, historical proposal conflicts and crash recovery. Preserve all 49 user tables and schema 25. Verify native event dependencies before declaring hidden controls vestigial.
