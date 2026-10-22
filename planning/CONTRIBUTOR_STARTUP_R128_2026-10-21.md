# C03 — explicit startup contributor

Delivered identity: r128 / 0.27.1-dev / schema 27. The fresh build, focused tests
and user-scoped release checks pass. Blank and synthetic Demo are uploaded and
their names, sizes and folder membership are verified. Access is closed.

## Behavior

An initialized workspace opens `frmContributorStart` with an empty selection on
every application launch. The user chooses an existing contributor or explicitly
adds a new one. Blank first-time setup already asks for a contributor name and
confirms that choice for the session. Both startup paths offer an Exit action.
The workspace context displays the confirmed contributor.

The retained `Workspace.ActiveProfileID` remains intact for historical data and
transfer compatibility. It alone no longer authorizes attribution. Session
TempVars bind the explicit selection to the current database path and workspace
GUID and survive a VBA reset. Missing or inconsistent session values fail closed.
Report saves, review decisions, comments/messages, operations, recovery ownership,
preferences, snippets and folder services use the central contributor guard.
Existing pending-editor/comparison guards still apply to contributor switching.

A successful production workspace transfer clears session selection. Reopening
asks from the imported contributor list. Historical directory backfill for older
schemas uses the imported source actor in a narrow synchronous compatibility
scope; it never confirms that actor as the current session's contributor.
There are no new tables, schema changes, deleted profiles or rewritten authors.

Contributor selection is attribution, not authentication or proof of identity.
Windows folder protections remain the user's accepted C01 boundary.

## Source and verification

The managed manifest includes `modContributorStartup` and the
`apply-contributor-startup` stage before finalization: 100 static modules,
four generated modules, 72 stages, and 67 verified final forms.
Build with the standard `scripts/develop.ps1`; do not assemble from a prior
release. `test-contributor-startup.ps1` uses disposable native copies and the
production Settings & Help importer. Existing populated-fixture automation makes
an explicit synthetic contributor choice rather than adding a product fallback.

The [startup procedure](../docs/Contributor_Startup.md) supplements the existing
user guide. C04 historical/effective-dated validation is the next critical item;
its policy rules are unchanged by C03. Physical display, keyboard/accessibility
and printed-fit acceptance remain separate.

## Evidence

The standard fresh build completed all 72 stages and verified 100 static modules,
67 forms, identity, retired routes and native unload guards. Build SHA256:
`55CC0C7C6B686A5F312FA52FC0DEA036D5FCA1F630B39F1ACDD39EAE266C8277`.
`build/workflow/20261021-114410-d6115faf-9566-4286-96e6-1b1fb8945e24/verification.json`
records PASS for BuildAndTest: 61 contributor assertions, import/reopen/cancel,
35 report-tools, 264 fixed-page/pending-buffer, 41 lifecycle, forced-close
save/recovery, smoke and 47 directory-merge checks. Blank user tables are empty.

**User-directed verification scope, 21 October:** the user explicitly waived the
old-output requirement for this delivery. The release run uses
`-CurrentFixturesOnly`; older schema fixtures and the populated schema-20 upgrade
are not retested. Previous release matrices remain historical evidence, not
current r128 compatibility proof. Current-schema production transfer, exact field
verification, source hashes, rollback, changed-source/populated-target and
unsupported-source guards remain in scope.

Final receipt:
`build/workflow/release-20261021-115911-fc9bd36e-f0da-4813-9e87-d246cc392928/verification.json`.
Both exact output files passed source/binding and startup/trust checks. Current
schema-27 merge history transferred exactly and reopened with guarded undo and
next-ID advancement. Full production import passed bootstrap rollback, changed
source, populated target, unknown table and future schema rejection. Additional
gates passed: directory/history/batch/UI, help/cleanup, Demo/report tools,
54 group-exit assertions, 16 adversarial exactness assertions, injected field
mismatch rollback/retry and out-of-order conversation/private-draft transfer.
The r127 Blank/Demo and restored r116 workspace match their recorded hashes.
See [complete verification](CONTRIBUTOR_STARTUP_R128_VERIFICATION.json) and
[current artifacts](../output/CURRENT_FILES.json) for hashes, counts and Drive IDs.

Each distribution is 18,219,008 bytes. Blank SHA256 is recorded above; Demo SHA256
is `5DC652E076093AE8E1426D5BFEFE384FC3556C2ABC3EC4B0BCD8248F2E37FB8E`.
Blank has zero user rows; Demo has 13 synthetic reports / five groups / 13 command
members and no scratch merge rows. Source changes are local; no Git commit/push
was performed. Unrelated folder renames and output cleanup were preserved.

Development probes passed 61 assertions plus reopen, Settings & Help transfer,
imported attribution and cancellation. The suite injects a workspace update
rejection to verify that adding a new contributor rolls back both the profile
and retained selection. Native domain lookups cannot see the uncommitted new
profile, so creation and selection use the same owned DAO transaction.

Earlier development attempts are not release candidates: nested `Eval` calls
exposed expected guard errors in the test harness; a diagnostic MCP import
doubled CRLF line endings; and a dynamic numeric COM setter rejected the setup
section height. Tests now call VBA services directly, MCP diagnostics normalize
line endings, and the generator uses `SetAccessProperty`. The final source-only
build and release gates must pass independently of those diagnostic copies.

Access Explorer's heuristic contributor-form check found no contrast, overlap,
out-of-bounds or truncation violations. The setup check found no new-button
overlap; it retained one pre-existing heading/help-button overlap warning.
These mechanical checks do not establish physical accessibility acceptance.
