# Standard build and smoke-test repair

19 October 2026. Repairs for the two verification blockers retained after priority batch four. No schema or report-data changes.

## Causes and changes

- The build scripts assigned computed .NET numbers directly to Access COM properties. The original autosave failure and a subsequent resource-tools failure showed incompatible Int16/Int32 setter casts; JSON layout values also arrived as Int64. A cached dynamic COM setter additionally confused string and Boolean property values. `access-properties.ps1` now converts values to the existing native property type and invokes Property.Value through reflection. The matching section/control-size calculations and authored-layout application use this common setter. Geometry formulas and all authored values are retained.
- The final event-binding step could create a duplicate `Form_Unload` beside an existing empty stub. This caused the later table-sort event-expression compilation to fail. The builder now reuses existing procedures. `ensure-native-unload.ps1` also normalizes previously generated empty duplicates, preserves the one nonempty procedure, and ensures the editor/grid guards are present. Multiple nonempty handlers are rejected instead of being silently discarded. The duplicate-stub fixture compiles after repair, and applying the repair twice retains one guarded procedure per form.
- `smoke-test.ps1` chained PowerShell COM access through `Forms.Item(...).Controls.Item(...)`. In a disposable copy, `frmLibrary` exposed 39 controls through native Access expressions while the PowerShell `Controls` lookup returned null. The original smoke run reproduced the failure at the Reports list inspection. The test now reads controls using Access expressions and writes values through a small test-only VBA helper. It continues to exercise the same product actions and expected results.
- Reports row counts now account for the actual ColumnHeads setting; empty lists are checked as empty. The smoke script's default filename now matches the standard builder's current default output filename.

- The fresh native build completed generation, compile and reopen checks, then its five-second final hash retry expired during Access/DAO shutdown. The builder now releases its script-owned COM references, collects finalizers, and permits a bounded 30-second retry. The exact revised cleanup/hash block is verified separately by `scripts/test-build-finalization.ps1` with retained DAO/form references. The original final-hash failure remains recorded in the log and verification JSON.

## Verification

The repaired broad smoke test passed all 48 assertions on the preserved batch-four candidate. This includes all three report types, revision history, restore/null values, archive guards, rollback injection, group filtering/moves, search with apostrophes, editor save, contextual help, and close/reopen persistence.

The disposable checkpoint build passed every remaining standard build stage, compilation and persisted unload-guard verification. Its rebuilt artifact also passed the 48-assertion smoke suite, 42 pending-cell assertions and 38 sorting assertions. The common property setter passed 348 authored-layout, repeat-application and recovery/event-binding checks. The independent standard build from blank also completed, compiled and retained both unload guards after reopening. That exact artifact passed the same 48 smoke, 42 pending-cell and 38 sorting assertions. Together with the 348 property checks, this is 476 counted assertions, plus the duplicate-stub repair/compilation checks.

Final blank artifact: `build/build-smoke-fixed.accdb`. Create-once inventory: `build/build-smoke-fixed.manifest.json`. Tracked evidence: `planning/BUILD_SMOKE_FIX_2026-10-19_VERIFICATION.json`.

The standard build and broad-smoke blockers are resolved. This is a local development build; retained-schema release transfer/rollback, NAVFIT compatibility, Access Runtime, physical display/accessibility and printed-fit acceptance remain separate gates. No release was published.

The previous batch-four candidate and populated r116 workspace retain their recorded hashes. No operational workspace has been upgraded or replaced.
