# Build pipeline

Current architecture and support decisions are in [PROJECT_STATE.md](../PROJECT_STATE.md).
This document owns generator ordering and build verification details.

For VS Code tasks, environment setup, the daily build-and-test command and
release verification, see [Development workflow](../docs/Development_Workflow.md).

The source of truth is `src/`, the ordered `scripts/build-manifest.json`, and the generator scripts it lists. Run a fresh full build with `scripts/build.ps1 -OutputPath <new file.accdb>`. Existing files are never overwritten. Identity comes from `src/modSwift.bas`; the schema must agree with `src/major-schema.json`.

## Ownership and order

The manifest currently lists 103 static VBA modules, four generated modules and 72 ordered stages across base resources, presentation and finalization. Full builds centrally load static source before generating forms and suppress the old repeated stage imports. The field catalogue retains one compatibility refresh of modUX alongside its generated metadata. Named field and surface catalogue generators replace execution of sliced script text. Incremental scripts retain their standalone imports and dependency behavior for disposable development copies.

### Canonical report fields

[field-definitions.json](../src/field-definitions.json) supplies report-column
types, recovery storage/control keys, field order and UI/validation metadata,
NAVFIT/PDF binding order, and fixed geometry. Native mapping/page wrappers read
the generated catalogue, rather than carrying separately generated VBA copies.
The three older Python mapping/page generators now refresh only compatibility
JSON views from this catalogue. They do not infer geometry or overwrite VBA.

To add a field, increment the catalogue version and append a stable ordinal with
its introduction `definitionVersion`, key, storage type, recovery key, control,
labels, applicability and validation metadata. Supply explicit per-form adapter
bindings and fixed-layout coordinates when the field participates in interchange
or the official editor. Preserve existing ordinals, captions, coordinates and
printed page associations. The first 79 fields retain the established service
facades; appended fields use the generic commit/restore/recovery path. Do not
extend official forms without supporting form/policy evidence.

Run `test-field-definitions.ps1` for exact existing binding/geometry checks; with
`-BuildPath <fresh blank>` it also exercises an appended synthetic field using
the production generators on a disposable catalogue/database, including a new
Access session's recovery. The report and recovery DDL, field/control arrays and
generic persistence loops consume the same metadata. New snapshots record the
catalogue version; prior snapshots retain their original definition version.

The fresh builder reasserts its canonical metadata before reopen and after final
stages, since historical standalone stages stamp their own older schema versions.
This does not upgrade an existing workspace in place or replace schema gates.

The consolidation audit retained these active layers. The 72 stages comprise one
base-resource stage, 65 presentation stages and six final stages. Fixed-page
geometry, visual decoration, unload-event installation and retired-route cleanup
have distinct responsibilities; do not drop a pass merely because it revisits
an existing form. Historical binary-template assemblers and the browser builder
are held in PROPOSED_PURGE, not alternative managed build paths.

`build-plan.ps1` validates files, root boundaries, inventory, stage dependencies, phase order and identity before Access opens. It fingerprints all declared inputs. `invoke-build-phase.ps1` executes stages in manifest order and records durations. Helpers deliberately share the build script scope because later generators consume catalogue/layout helpers.

`finalize-application.ps1` owns final legacy-route cleanup after presentation decorators. It removes obsolete Reports folder controls, labels the still-active organizational folder tools, redirects old help bindings, and rejects unresolved retired form references before deletion. Folder data, moves, history and undo remain intact. `frmMajorWorkbench` remains a required context host. Hidden `lstFolders` is a compatibility name for the summary-group filter, not a folder ID.

After compile and reopen, `verify-built-application.ps1` compares installed VBA with all static source modules, checks generated modules, metadata, native editor/group unload guards, retired form absence and event function bindings. Input fingerprints must still match. The adjacent `.accdb.build.json` receipt contains artifact hash, identity, input fingerprints, stage durations and structural results. A receipt is not a substitute for workflow, migration or physical acceptance tests.

## Verification

C10 removes duplicate base help/checklist/comments construction, the old note
composer and the temporary review screen. The review-return stage now creates
frmReportReviewPanel directly. Decoration and final binding checks retain their
order. Rails and compact selectors use the six routes from modUX rather than
parallel destination lists. Required group/folder hosts remain. See the
[producer/consumer and generation audit](CONSOLIDATION_PLAN.md#c08c10-producer-consumer-and-navigation-audit).

Run `scripts/test-build-plan.ps1` without Access to exercise rejected plans and input drift. Run `scripts/test-build-cleanup.ps1 -BuildPath <fresh blank>` for native legacy adapters, folder history and pending report navigation. Select other suites for the changed behavior; see [contributing](../CONTRIBUTING.md). Release preservation requirements and the recorded historical-test exception are described in [AGENTS.md](../AGENTS.md) and [PROJECT_STATE.md](../PROJECT_STATE.md).

Add a generator to the manifest rather than adding a hidden nested full-build call. Declare its inputs, dependencies and phase. Any source/input change during a build invalidates its receipt; rebuild from the final source. Keep old release binaries and populated workspaces untouched. Publish only verified Blank and synthetic Demo copies in `output/`.
