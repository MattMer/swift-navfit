# Performance improvements r100 - 2026-10-10

App 0.25.0-dev; schema 25. New candidates are based on r99. Published r98 and existing workspaces remain untouched.

## Implemented

- Unchanged group navigation reuses a comparison whose snapshot stamp matches the current group. Comparisons with any proposals are excluded. Changed revisions, membership and workflow stamps require a fresh baseline. Historical snapshot rows remain immutable.
- Live editor validation caches findings by report identity, exact pending values (length-prefixed, with separate Null markers), and calendar date. Changed input still validates immediately. Dirty state is repainted independently. Explicit validation summaries continue to evaluate current inputs.
- Recovery checkpoints skip writes when the pending values, session, saved baseline and contributor match the previous successful checkpoint. Deleted checkpoint rows are recreated; failed writes never update the fingerprint. Clean idle checks skip unnecessary deletes and Overview refreshes.
- Saved group averages in the editor read the report's saved trait columns in one joined query, replacing seven revision-value queries per member. Pending ITA remains separate.
- Comparison average refresh writes an ITA cache row only when its displayed average actually changes.

## Measurement

Three unchanged opens/closes of a three-member synthetic group:

| Measure | r98 | r100 |
|---|---:|---:|
| GroupSnapshot | 3 | 1 |
| GroupSnapshotMember | 9 | 3 |
| ComparisonSet | 3 | 1 |
| ComparisonMember | 9 | 3 |
| Total added history rows | 24 | 8 |

67% fewer history rows; later unchanged opens add none. Evidence: PERFORMANCE_R100_NAVIGATION_EVIDENCE.json and REVIEW_R98_NAVIGATION_EVIDENCE.json.

Timing did not demonstrate a speedup: r98 opens were 853-971 ms; r100 opens were 998-1128 ms. These are separate small-fixture local runs, not a controlled scale benchmark. No speed claim is made.

## Verification

- Compiled blank, synthetic demo and preserved test workspace candidates.
- 16 new baseline reuse, changed revision, immutable history, proposal isolation, checkpoint skip/recreation and live validation assertions.
- 42 pending-cell failure, partial-save, retry, discard, close cancellation and attribution assertions.
- 26 approved-trait rollback/retry assertions.
- Queued cell identity, coalescing and failed-save retry suite.
- Recovery suite: 12 preparation and 23 recovery assertions, new Access session and automatic active-text timer.
- Demo normal startup opens Overview.
- All 49 user tables are identical between r99 and r100 preserved workspaces.

## Remaining optimization opportunities

Successful individual comparison cell saves still create fresh immutable group/comparison baselines. Removing these requires separating the live Members view from durable proposal baselines and testing all historical proposal paths; this change deliberately preserves those guards.

Comparison averages still scan the cache; they avoid redundant writes but do not yet maintain incremental group totals. The group selector still refreshes during editor updates. Native display and keyboard review and larger-group performance measurements remain outstanding. No GitHub publication performed.

Final author UI suite: 354 properties, obsolete-control absence, startup and populated navigation checks passed. Candidate SHA256 hashes are recorded in PERFORMANCE_R100_VERIFICATION.json.
