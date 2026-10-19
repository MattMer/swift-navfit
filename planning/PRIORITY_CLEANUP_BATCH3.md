# Priority cleanup - implementation batch 3

19 October 2026. Group-calculation consolidation from the current-state review. Local development candidate; no release published.

## Implemented

- Group average reads current group membership and all seven saved revision traits in two queries, replacing seven revision-value queries per included member. It preserves numeric-mark acceptance, whole-report NOB and archive exclusions, Null results, and the missing-historical-field error.
- PSG RSCA reads committed Report fields and archive flags together instead of calling SavedFields and a separate archive lookup for each member. Existing senior identity, paygrade, exclusion confirmation, exact/rounded baseline, zero-baseline and partial-trait rules remain intact.
- The group calculation explanation uses its already loaded committed Report rows. The calculator shares one read-only member snapshot between PSG and explanation; each consumer clones it without moving the caller's cursor, including when the caller is at EOF.
- The group workbench passes its calculated PSG result to the footer instead of recalculating it. Other callers continue to calculate fresh results by default.
- There is no long-lived calculation cache, schema change or authoritative-data write. Pending editor/grid buffers and fixed report geometry are untouched.

These inputs are intentionally distinct: group average uses revision trait values; PSG and the explanatory context retain their prior committed Report-field semantics. The tests deliberately diverge Report traits from revision values to verify that this distinction survives. Individual ITA remains a separate calculation.

## Verification

401 native checks passed:

| Suite | Checks |
|---|---:|
| Exact legacy/new group results, explanations, cursor independence, historical-field errors, no history writes and UI output | 204 |
| Existing group correctness | 39 |
| Existing reporting-senior/RSCA/comparison workflows | 39 |
| Existing contextual resource tools | 119 |

The initial test-harness compile error shown during execution was caused by a variable named `shared`. It was renamed before the successful run. The harness and preserved legacy reference functions exist only in disposable test copies, not in the candidate.

All six changed standard modules matched exports from the candidate after normalizing Access formatting while preserving string contents. Compilation passed. The create-once final manifest confirms all user-data tables are empty.

## Native calculation timings

Combined group average + PSG RSCA + explanation, including fresh member loading. Five paired samples per size, alternating old/new execution order. Every complete value/flag/reason/explanation comparison matched, and reads created no report revisions or historical snapshots.

| Members | Legacy median | New median | Legacy maximum | New maximum |
|---:|---:|---:|---:|---:|
| 25 | 225 ms | 9 ms | 267 ms | 10 ms |
| 50 | 462 ms | 13 ms | 510 ms | 16 ms |
| 100 | 938 ms | 21 ms | 963 ms | 22 ms |

These are synthetic native calculation timings, not screen paint, cold-start guarantees or production tail latency. All 15 paired samples are retained in `build/batch3-group-scale.csv`.

## Artifacts and limits

- Candidate: `build/priority-batch3-final.accdb`.
- Manifest: `build/priority-batch3-final.manifest.json`.
- Tracked verification: `planning/PRIORITY_CLEANUP_BATCH3_VERIFICATION.json`.
- Builder: `scripts/build-ux-group-calculations.ps1`; copies the previous candidate, refuses overwrites and imports the six related modules together.
- Native test and measurement entry points: `scripts/test-ux-group-reads.ps1` and `scripts/measure-ux-group-reads.ps1`.

The batch-two candidate and original populated r116 workspace remain unchanged. This candidate uses the separate template-update build path. The earlier full-build comparison-autosave type-cast failure and broader smoke-test Reports-inspection failure remain outstanding; this batch does not claim those gates passed. Full retained-schema migration/rollback acceptance, physical display/accessibility and printed fit remain release requirements.

Remaining review work includes member-grid context batching, My Work inclusion/count queries, in-app build/workspace identity, and the save/shared-field/contextual-action presentation improvements.
