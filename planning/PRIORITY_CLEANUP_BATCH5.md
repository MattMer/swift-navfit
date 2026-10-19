# Priority cleanup - implementation batch 5

19 October 2026. Finishes My Work inclusion/count batching for P02/P03 in the current-state improvement backlog. Separate local development candidate; no release published.

## Implemented

- `UXWorkVisibleRows` creates fresh visible work keys from batched report participation, thread authors/events, route steps, assignment/report and route/report reads. It replaces calls to the single-item database visibility service for every candidate row.
- Active queue preparation consumes those keys and retains the existing display materialization, text, sorting and selection behavior. History continues to include completed work and all existing contributors according to its prior rules.
- Overview counting uses the same fresh inclusion service without formatting rows or replacing the display snapshot. Opening a history view or another contributor's view is unaffected by counting the current contributor's work.
- Contributor visibility retains the separate assignment, ordered-route and comment-thread semantics: reply authors, saved-report contributors, active task participants, current route assignees, requesters and final approvers. Thread resolution, archived/stale routes, group membership stamps and accepted-route behavior are preserved.
- The original `UXWorkVisible` and authoritative review/write services remain intact. No schema, form layout, pending-value, recovery or persistent-cache change was made.

The inclusion service uses seven thin recordsets and reuses the existing group-stamp result once per distinct referenced pending/accepted route group. Queue presentation retains its separate refresh-scoped stamp cache. This removes per-candidate visibility database calls; it does not make total cost independent of workspace history or distinct groups.

## Verification

2,839 counted assertions passed:

| Suite | Assertions |
|---|---:|
| Exact My Work rows/order, contributor scope, history/count independence, navigation, actions and layout | 2,487 |
| Paired queue/count equivalence, read-only checks, case semantics and missing-revision errors | 32 |
| Callback-free inclusion/count/display checks and active rejection probe | 3 |
| Queue refresh ownership | 7 |
| Authoritative routing workflows | 38 |
| Review comments | 28 |
| Shared Reports module: fields, historical values, filters and sorting | 196 |
| Broad smoke workflows | 48 |

The scope fixture includes an unrelated third contributor and reply-only participation. Calling Overview count immediately after preparing each contributor/history view must leave every prepared cell and order unchanged. The edge fixture checks that binary VBA state/event comparisons and the existing SQL participation comparisons keep their distinct legacy case behavior; a missing saved revision still raises the original error.

A disposable test replaces `UXWorkVisible` with an error-raising rejection probe, confirms that probe is active, then successfully performs fresh inclusion, count and prepared queue reads. The probe and legacy reference procedures are absent from the candidate. Scale reads are also checked for unwanted revision/snapshot writes.

Both changed modules (`modUXReadRows`, `modUXMyWork`) matched exports from the final candidate after Access-format normalization with string contents preserved. Compilation passed. The create-once manifest confirms all user-data tables are empty. The earlier full-build repair candidate, batch-four candidate and populated r116 workspace retain their recorded hashes.

## Native measurements

The reference is the already optimized batch-four implementation, not the original per-cell formatting path. Synthetic groups contain one assignment, one group route and one root comment per report. Three paired samples per size, alternating legacy/new order; every queue value/order and count matched. Medians:

| Reports | Area | Previous batch | Batch five |
|---:|---|---:|---:|
| 25 | Queue | 129 ms | 23 ms |
| 25 | Count | 109 ms | 4 ms |
| 100 | Queue | 539 ms | 59 ms |
| 100 | Count | 500 ms | 8 ms |

The 100-report fixture contains 300 work items. These measurements cover native queue preparation/materialization and counting; they are not physical screen-paint, cold-start or production tail-latency guarantees. Raw pairs are retained in `build/batch5-work-inclusion-scale.csv` and copied into tracked verification JSON.

## Artifacts and limits

- Candidate: `build/priority-batch5-final.accdb`.
- Manifest: `build/priority-batch5-final.manifest.json`.
- Tracked verification: `planning/PRIORITY_CLEANUP_BATCH5_VERIFICATION.json`.
- Builder: `scripts/build-ux-work-inclusion.ps1`; copies the prior blank artifact, refuses overwrites and imports both related modules together.
- Native comparison/measurement: `scripts/test-ux-work-inclusion.ps1`; test-only legacy procedures were exported from the preserved pre-batch-five artifact.

The prior [standard build/smoke repairs](BUILD_SMOKE_FIX_2026-10-19.md) are retained. This batch updates that compiled blank template through a module-only build and runs the broad smoke suite; it does not repeat full form generation or the retained-schema release transfer/rollback matrix. Physical display/accessibility, Access Runtime, NAVFIT exchange and exact printed fit remain separate acceptance gates.

Remaining priorities include consistent save/recovery status, visible shared-field scope and affected approvals, contextual next actions, build/workspace identity, and membership/directory read cleanup.
