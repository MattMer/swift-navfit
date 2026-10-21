# Build pipeline

The source of truth is `src/`, the ordered `scripts/build-manifest.json`, and the generator scripts it lists. Run a fresh full build with `scripts/build.ps1 -OutputPath <new file.accdb>`. Existing files are never overwritten. Identity comes from `src/modSwift.bas`; the schema must agree with `src/major-schema.json`.

## Ownership and order

The manifest lists 94 static VBA modules, four generated modules and 70 ordered stages across base resources, presentation and finalization. Full builds centrally load static source before generating forms and suppress the old repeated stage imports. The field catalogue retains one compatibility refresh of modUX alongside its generated metadata. Named field and surface catalogue generators replace execution of sliced script text. Incremental scripts retain their standalone imports and dependency behavior for disposable development copies.

`build-plan.ps1` validates files, root boundaries, inventory, stage dependencies, phase order and identity before Access opens. It fingerprints all declared inputs. `invoke-build-phase.ps1` executes stages in manifest order and records durations. Helpers deliberately share the build script scope because later generators consume catalogue/layout helpers.

`finalize-application.ps1` owns final legacy-route cleanup after presentation decorators. It removes obsolete Reports folder controls, labels the still-active organizational folder tools, redirects old help bindings, and rejects unresolved retired form references before deletion. Folder data, moves, history and undo remain intact. `frmMajorWorkbench` remains a required context host. Hidden `lstFolders` is a compatibility name for the summary-group filter, not a folder ID.

After compile and reopen, `verify-built-application.ps1` compares installed VBA with all static source modules, checks generated modules, metadata, native editor/group unload guards, retired form absence and event function bindings. Input fingerprints must still match. The adjacent `.accdb.build.json` receipt contains artifact hash, identity, input fingerprints, stage durations and structural results. A receipt is not a substitute for workflow, migration or physical acceptance tests.

## Verification

Run `scripts/test-build-plan.ps1` without Access to exercise rejected plans and input drift. Run `scripts/test-build-cleanup.ps1 -BuildPath <fresh blank>` for native legacy adapters, folder history and pending report navigation. Run smoke, report tools, My Work, fixed pages and lifecycle suites on disposable copies. Retained-schema migration, populated transfer, exactness and rollback gates remain release requirements.

Add a generator to the manifest rather than adding a hidden nested full-build call. Declare its inputs, dependencies and phase. Any source/input change during a build invalidates its receipt; rebuild from the final source. Keep old release binaries and populated workspaces untouched. Publish only verified Blank and synthetic Demo copies in `output/`.
