# Priority cleanup - implementation batch 2

Implemented My Work row preparation and Reports batch reads from the current-state review. This is a local development candidate, not a published release.

## Changes

- My Work joins each work-item table to report context and builds each included display row once. Contributor names, route-step counts/assignees and latest comment-thread events are loaded in sets. List cells read an in-memory display row rather than reopening records for each column. The original query remains available for exact comparison and compatibility callers.
- Reports loads report/revision metadata, saved trait values, group titles, assignment facts and route facts in sets. Names use already selected report fields. ITA retains the original saved-revision rules: missing fields mean Unavailable; null/empty/NOB traits are excluded; report NOB overrides the average. Current Report trait columns are not substituted for revision values.
- Status filtering reads the same prepared status shown in the row. Group-label order, revision numbering, column layout, selection and sorting remain intact.
- Group stamps are reused within each display preparation. Maps are published only after a complete refresh and rebuilt on each normal list refresh. No new tables, schema changes or authoritative data writes are introduced.
- Accept/return/approve services and their live freshness guards are unchanged. My Work inclusion still uses the existing visibility predicate once per candidate item; set-based inclusion/count optimization remains future work. These maps are display snapshots, not authorization or write baselines.
- Full and relevant incremental builders load the new module. `build-ux-read-rows.ps1` creates a separate candidate and refuses overwrites.

## Verification

1,951 targeted native checks passed:

| Suite | Checks |
|---|---:|
| My Work scope, all cell values/order, history, contributor visibility, returned/cancelled assignments, stale/group-stale/archived routes | 1,672 |
| Reports historical traits, state precedence, group labels, filters, selection and sorting | 196 |
| Refresh ownership, including retained activation refresh | 7 |
| Routing and decision guards | 38 |
| Query/value-list/header/grid sorting | 38 |

All four changed/new standard modules matched their candidate exports after normalizing Access formatting while preserving string contents. Candidate compilation passed. The final read-only inventory found every user-data table empty and recorded the exact hash. Existing workspace files and the first-batch template were preserved.

## Measured query preparation

Five paired samples per fixture/view, alternating the execution order, on native Access using disposable synthetic copies. Timing includes building the prepared maps and reading every result field. All 20 complete result comparisons matched exactly; query reads created no report revisions or group/comparison snapshots.

| Reports in fixture | View | Legacy median | Prepared median | Legacy max | Prepared max |
|---:|---|---:|---:|---:|---:|
| 25 | Reports | 500.0 ms | 15.6 ms | 1,656.3 ms | 23.4 ms |
| 25 | My Work | 546.9 ms | 62.5 ms | 601.6 ms | 70.3 ms |
| 100 | Reports | 2,093.8 ms | 62.5 ms | 2,320.3 ms | 70.3 ms |
| 100 | My Work | 2,257.8 ms | 226.6 ms | 2,320.3 ms | 250.0 ms |

Fixtures include one assignment per report, a group-linked route every fifth report and a comment every third report. Membership changes leave earlier group routes stale, exercising that path. These are query/data-preparation timings, not physical screen paint, cold-start guarantees or production p95. Five samples do not establish tail latency. First samples remain individually recorded in `build/batch2-read-scale.csv`.

## Candidate and remaining work

- Candidate: `build/priority-batch2-final.accdb`, based on the preserved first-batch blank.
- Create-once artifact inventory: `build/priority-batch2-final.manifest.json`.
- Tracked summary: `planning/PRIORITY_CLEANUP_BATCH2_VERIFICATION.json`.
- Raw native logs and source exports: `build/batch2-*`.

The previous from-source build failure in the comparison-autosave section-height assignment and the broader smoke-test Reports-inspection failure were not repaired or claimed passing in this batch. The candidate was built through the isolated template-update path. Full release compatibility/rollback checks, physical display/accessibility and printed fit remain release gates.

Further batches can address set-based My Work visibility/counts, filtering fact reads to smaller library subsets, shared group calculations/member-grid context, in-app build/workspace identity, and save/shared-field/action presentation. No release was published or operational database updated.
