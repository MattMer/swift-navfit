# Build and legacy cleanup - r125

The full builder now uses an ordered manifest with 70 stages, 94 static VBA modules and four generated modules. Preflight rejects missing dependencies, invalid ordering, schema/identity disagreement, inventory drift and paths outside the repository before Access opens. A central static-module inventory replaces repeated stage imports; the field catalogue retains one compatibility refresh of modUX. Named catalogue generators replace execution of sliced script text; standalone incremental scripts retain their imports and dependency calls.

The base database reference has a distinct owner and is released immediately after schema creation. This fixes the old final-close failure caused by a child script reusing the shared `$db` variable. Cleanup releases owned COM references without enumerating COM collections and tolerates already-released aliases; a focused nested-array/alias test passes. Finalization owns native unload installation, stale route cleanup and reopened verification. The source comparison normalizes Access-added empty `Call` parentheses and code formatting while preserving literal strings exactly. Build receipts fingerprint declared inputs, record stage times and artifact identity, and reject input changes during generation.

## Supported and retired routes

Organizational folder tools are still active and retained, with clearer labels. Folder creation, hierarchy, moves, batch history and guarded undo remain available. Obsolete Reports controls are removed. Old AddFolderUI/MoveFolderUI callbacks open the supported folder tools instead of interpreting the summary-group filter as a folder ID. MoveReportUI uses the summary-group preview. Main-menu compatibility callbacks use normal pending-save/group/senior exit guards.

The duplicate `frmHelp` is retired after event/subform inspection; its callback redirects to `frmUXHelp`. Previously retired standalone report comments, review, checklist and note-composer windows remain absent. The hidden `frmMajorWorkbench` context host remains necessary and supported. Hidden `lstFolders` retains its group-filter compatibility identity.

## Documentation

The roadmap now describes current scope and remaining work; the old roadmap is archived intact. Contributor guidance describes manifest ownership, source verification and current acceptance suites. Historical UI guides are marked with links to the current user guide and current-state record. Historical review-comment/integrated-tools harnesses still contain superseded standalone UI and are not release acceptance suites; focused current tests cover their retained services.

## Verification scope

Native testing uses disposable synthetic copies and preserves source artifact hashes. Physical visual/DPI, keyboard/screen-reader and printed-layout acceptance remain deferred. Full NAVFIT98A interchange certification is not implied. See the adjacent verification JSON for final artifact hashes, test evidence and distribution readback.
