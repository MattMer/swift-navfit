# Priority cleanup - implementation batch 1

Scope: first measurable cleanup from the 18 October current-state review. This is a development candidate, not a published release.

## Changes

- Give My Work initialization one unified-queue refresh owner. Preserve the default refresh for contributor changes and review actions; avoid a second load on return navigation and the legacy opening path.
- Count visible work directly using the same inclusion predicate as the list, without evaluating report labels, cell display values or sorting. Visibility/freshness checks remain authoritative and unchanged.
- Keep the legacy task list: assignment decisions still use its selection.
- Correct the README's r113 download reference and distinguish historical release verification from current local files.
- Add a create-once JSON artifact inventory with schema, app version, build revision, source checkpoint/file hashes, artifact hash and all non-system/non-metadata table counts. It reads through read-only DAO, rejects linked tables, verifies the file hash remains unchanged, and refuses blank certification for any populated table. It does not infer synthetic provenance or claim installed-source alignment/release acceptance.

## Artifact protection

The local r116 and r115 files named Blank are populated and were preserved. The retained r114 blank passed the all-table empty check. The separate candidate uses that template plus the existing r115 compatibility and r116 report-rail builders, followed by the three changed modules. No schema change is introduced.

The full from-source build failed in the pre-existing `apply-ux-comparison-autosave.ps1:14` section-height assignment with `Specified cast is not valid`. The partial build is not a usable deliverable. The staged-template candidate is separate; full-build repair remains outstanding.

## Remaining review work

Single-pass row materialization for My Work and Reports, shared calculation contexts, in-app build/workspace identity, unified save/recovery presentation, shared-field scope and contextual actions remain subsequent batches. A new release still requires the retained-schema migration matrix and the other release gates in AGENTS.md. Physical display/accessibility and printed fit remain unverified.

## Verification

- Candidate compiles in native Access.
- 38 My Work assertions passed, including count equivalence with formatted rows across state changes.
- Seven instrumented refresh-path checks passed on a disposable copy. Actual forms also refresh on activation through the sorting wrapper's preserved `UXWorkChanged` binding. Initialization now loads once; a newly opened form still performs its separate activation refresh. That freshness behavior is retained deliberately.
- 38 routing assertions passed, including returned tasks, stale revisions and archived-report guards.
- Final candidate passed all-table empty validation; hash and source inventory are in `build/priority-cleanup-candidate.manifest.json`.
- The populated local r116 failed RequireBlank as expected and retained its hash. The retained r114 passed; r115 was also found populated.
- No new performance percentage is claimed; representative timing and the full release matrix remain pending.

The broader `smoke-test.ps1` passed its data-integrity checks (including no-op revision behavior and transaction rollback) and opened Home/Reports, then stopped at line 85 with a null-valued expression in the Reports list inspection. It did not complete and is not counted as a passing suite. This and the full-build layout failure require follow-up before release.

Targeted passing total: 83 assertions/checks (38 My Work + 7 refresh ownership + 38 routing). Candidate path: `build/priority-cleanup-candidate.accdb`. Existing operational files were not updated; no release was published.
