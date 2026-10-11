# Live Members and action-bound guards - r101 - 2026-10-10

Implemented single-user efficiency priorities 1 and 2. App 0.25.0-dev, schema 25. Candidates derive from r100; previous databases are untouched.

## Changes

- Members opens with a live group context, loads current saved Report values and GroupMember membership, and leaves txtMeeting at zero. No GroupSnapshot, GroupSnapshotMember, ComparisonSet or ComparisonMember writes occur.
- Direct trait, promotion and career-recommendation saves create normal atomic report revisions, update the cache and loaded state, and do not copy the group into historical baselines.
- Explicit historical Comparison entry still creates/reuses a comparison baseline. Staging proposals from Members materializes a baseline only when requested; existing historical proposals and apply conflict checks remain intact. A changed selected report revision is rejected before staging.
- Row display/focus uses an in-memory dictionary populated from joined report archive/workflow/revision state. UXGridLockReason's display branch has no database queries and no GroupStamp evaluation.
- Save uses an authoritative report read, current-revision comparison, archive/workflow checks and group membership/archive checks. Historical comparison saves retain group stamp checks. These protect against other local forms/imports as well as unintended external changes.
- Pending queue attribution, partial-save retry/discard, approval reopen/rollback, report revision services and historical proposal services remain intact.

## Evidence

Three opens/closes of the three-member synthetic group add zero rows to all four group-history tables (r98: 24, r100: 8). See LIVE_MEMBERS_R101_NAVIGATION_EVIDENCE.json. Direct trait/career saves also add no group snapshots in the new tests across EVAL, CHIEFEVAL and FITREP.

Observed opens: 999, 1067 and 1081 ms. No navigation speedup claim: tiny fixture and separate local runs. Eliminated writes and display-time queries are the established improvements.

## Verification

- All three r101 candidates compile.
- 61 new assertions: all report types, live opening/saves, zero history writes, latest-value reopening, stale saved revision, removed membership, explicit historical baseline, stale proposal rejection and lazy proposal staging.
- 42 pending-cell failure/retry/discard/attribution/native-close assertions.
- 26 approved-trait rollback/retry/repeated-save assertions.
- Repeated native control events across all three types.
- Queue row identity/coalescing/preservation/retry.
- 354 author properties/startup/populated navigation assertions.
- All 49 user tables identical between preserved r100 and r101 workspaces.

## Limits and next work

Single-user exclusive opening is not enforced by this change. Write-bound guards remain; this work does not rely on removing them. Display state can remain cached after another form changes a report, but an attempted save is rejected and reopening reads current data.

Editor refresh targeting, legacy comparison list/control removal, group navigation restructuring and idle-timer event-driven changes remain priorities 3-6. Explicit historical comparison edits retain the historical baseline-copy behavior; ordinary Members direct edits no longer use that path. No schema changes or GitHub publication. Native physical display/accessibility review remains outstanding.
