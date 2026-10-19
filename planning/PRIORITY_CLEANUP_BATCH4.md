# Priority cleanup - implementation batch 4

19 October 2026. Member-grid read batching from the current-state review. Local development candidate; no release published.

## Implemented

- Member-grid refresh loads confirmed identity and prior-history candidates together, replacing separate identity/date/history lookups per row. Selected ITA keeps the existing source-specific calculation and is reused for the same source during that refresh only. Equal-ending-date ties delegate to the existing ambiguity logic, preserving its exact result and period selection.
- Saved revision labels and historical comparison fields load together. Live Members retains committed Report-field semantics; historical comparison retains its captured RevisionValue semantics, including Null values and the missing-field error.
- Historical proposed changes load in one ordered query and retain proposal-order overlays. Live Members skips the proposal query because it has no comparison set.
- The temporary grid is populated in one transaction in its separate DAO workspace. A failed refresh rolls back, unbinds and closes the temporary cache, then raises the original error. Error cleanup exits the active handler before touching invalidated recordsets so cleanup cannot mask the cause.
- Pending-cell refusal remains before load cleanup. Retry/Discard, authoritative save-time revision/membership checks, report saves, recovery and fixed report geometry retain their existing behavior. No schema change or persistent calculation cache was added.

## Verification

7,603 counted native assertions passed, plus the existing trait-queue suite (which does not report a numeric count):

| Suite | Assertions |
|---|---:|
| Exact grid fields/order, live/historical proposal overlays, read-only history and failed-load recovery | 7,390 |
| Directory/history identity, original/batched prior values, NOB/incomplete/ties/archive and date boundaries | 72 |
| Pending cells, partial save, retry, database failure, close cancellation, discard and attribution | 42 |
| Live navigation/save, authoritative guards, three report types, historical conflict and lazy proposals | 61 |
| Sorting/query/header/selection/filter behavior | 38 |

The equivalence fixture compares all 21 grid columns at 25, 50 and 100 members in live and historical modes. It removes a saved historical field after a partial cache build, confirms the original missing-field error and unbound form, restores the fixture field, and verifies reload succeeds. Original and test-only legacy loads run in disposable copies. The legacy load is appended only to the test module; it is absent from the candidate.

The older prior-grid test expects career controls to lock immediately after an external editor save. It fails identically on the unchanged batch-three candidate. This batch leaves that existing behavior unchanged; the current LiveMembers suite confirms row focus uses cached state and stale writes are rejected at the authoritative save boundary. The older test is not counted as passing.

The missing-comparison-form dialog encountered during testing came from an unconditional test-runner Discard cleanup after the form had closed. The runner now checks whether the form is loaded before invoking cleanup. The final complete run finished without that cleanup error.

Compilation passed. Both changed modules matched candidate exports after normalizing Access formatting while preserving strings. The create-once manifest confirms all user-data tables are empty. The prior candidate and populated local r116 workspace still match their earlier inventory hashes.

## Native load samples

One legacy/new pair per size and mode from the final successful run, legacy first. Each member has confirmed identity and an external prior evaluation; historical views include proposed edits. Timings include fresh temporary-grid construction and binding, not physical screen-paint completion.

| Members | Live legacy | Live new | Historical legacy | Historical new |
|---:|---:|---:|---:|---:|
| 25 | 264 ms | 176 ms | 322 ms | 136 ms |
| 50 | 385 ms | 293 ms | 693 ms | 397 ms |
| 100 | 948 ms | 417 ms | 3,045 ms | 1,205 ms |

These are synthetic samples, not medians, production tail latency or cold-start guarantees. Raw results are retained in `build/batch4-grid-scale.csv` and copied into the tracked verification JSON. Source-specific ITA reads still occur for each distinct selected prior source; deeply populated member histories may cost more than this fixture.

## Artifacts and limits

- Candidate: `build/priority-batch4-verified.accdb`.
- Manifest: `build/priority-batch4-verified.manifest.json`.
- Tracked verification: `planning/PRIORITY_CLEANUP_BATCH4_VERIFICATION.json`.
- Builder: `scripts/build-ux-grid-reads.ps1`; copies the prior candidate, refuses overwrites and imports both related modules together.
- Native comparison and measurement entry point: `scripts/test-ux-grid-reads.ps1`.

Earlier intermediate batch-four files are diagnostic candidates; use the verified artifact above for review. This is a separate blank development candidate. No workspace was upgraded or replaced, and no deployment or release was published.

The previous full-build comparison-autosave type-cast failure and broad-smoke Reports-inspection failure remain outstanding. Full retained-schema transfer/rollback, blank/demo release acceptance, physical display/accessibility and printed-fit checks remain release requirements.

Remaining review work includes My Work inclusion/count queries, in-app build/workspace identity, and save/shared-field/contextual-action presentation improvements.
