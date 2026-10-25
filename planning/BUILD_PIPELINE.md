# Build pipeline

Current architecture and support decisions are in [PROJECT_STATE.md](../PROJECT_STATE.md).
This document owns build ordering and build verification details.

For VS Code tasks, environment setup, the daily build-and-test command and
release verification, see [Development workflow](../docs/Development_Workflow.md).

The source of truth is `src/` (VBA modules, catalogues and the forms in
`src/forms`), `scripts/build-manifest.json` and the resource scripts it lists. Run
a fresh build with `scripts/build.ps1 -OutputPath <new file.accdb>` (about two
minutes). Existing files are never overwritten. Identity comes from
`src/modSwift.bas`; the schema must agree with `src/major-schema.json`.

## Order

1. Create the database: tables, validation rules, relationships and canonical
   queries from `major-schema.json` and `build.ps1`.
2. Load every static module in `sourceModules`, then generate `modTraitCatalogue`.
3. Run the `resources` stages in manifest order. Each generates something that
   is not a form:
   - `apply-pdf-print-catalogue`: `modPDFPrintCatalogue` from `pdf-print-layout.json`
     (checks the embedded template image hashes).
   - `apply-ux-field-catalogue`: `modUXFields` from `field-definitions.json`, which
     carries runtime field layout, bindings and labels; refreshes `modUX`.
   - `apply-ux-surface-catalogue`: `modUXSurfaceFields`.
   - `resource-navfit-template`: `modNAVFITTemplate`, embedding the reviewed
     native blank.
   - `resource-pdf-report`: the `rptPDFOutput` report with the eight template
     images and its text slots. It stays generated, not stored as text, because
     its printer settings differ on every save.
   - `resource-startup-properties`: hides the navigation pane.
4. Load every form in `forms` from `src/forms` (`install-form-text.ps1`).
5. Restore identity metadata, load `AutoExec`, compile, reopen and run
   `verify-built-application.ps1`.

## Forms are source

Each form is stored as normalized Access text in `src/forms/<name>.form.txt`,
including its class module. Normalization drops only per-build noise that
`LoadFromText` does not need: the checksum, generated GUID, name-map and
record-source-date blocks, and the `BaseInfo` combo cache. The stored text is a
fixed point: building it and exporting again gives the same files.

- Small changes: edit the text directly.
- Layout work: build a disposable copy, change the form in Access design view,
  then run `scripts/export-form-text.ps1 -BuildPath <that copy> -Form <name>`. It
  writes normalized text into `src/forms`. Add a new form to `forms` in the manifest.
- Catalogue changes include the form edit. `field-definitions.json` drives the
  generated runtime layout, but each field's control must exist on `frmEditor`
  and `frmUXEditorFields` (the report type is the one field the surface does not
  edit). `test-field-definitions.ps1` fails without it.

## Canonical report fields

[field-definitions.json](../src/field-definitions.json) supplies report-column
types, recovery storage and control keys, field order, UI and validation
metadata, NAVFIT/PDF binding order and fixed geometry. The three Python
mapping/page generators refresh only the compatibility JSON views from it.

To add a field, increment the catalogue version and append a stable ordinal with
its introduction `definitionVersion`, key, storage type, recovery key, control,
labels, applicability and validation metadata. Supply explicit per-form adapter
bindings and fixed-layout coordinates when the field participates in interchange
or the official editor, and add its control to the two editor forms. Preserve
existing ordinals, captions, coordinates and printed page associations. Do not
extend official forms without supporting form or policy evidence.

## Verification

`build-plan.ps1` validates files, root boundaries, source-module and form
inventories, stage order and identity before Access opens, and fingerprints all
declared inputs, including every form file. After compile and reopen,
`verify-built-application.ps1` compares installed VBA with the static source
modules and checks generated modules, metadata, native unload guards, retired
form absence and event function bindings. The adjacent `.accdb.build.json`
receipt records the artifact hash, identity, input fingerprints, stage durations
and structural results; it does not replace workflow, migration or physical
acceptance tests.

Run `scripts/test-build-plan.ps1` without Access for rejected plans, input
drift, form normalization and the runner inventory, and
`scripts/test-field-definitions.ps1` for catalogue and form-control agreement.
Select other suites for the changed behavior; see [contributing](../CONTRIBUTING.md).

Add a non-form resource as a manifest stage rather than a hidden nested build
call, and declare its inputs. Any source or input change during a build
invalidates its receipt; rebuild from the final source. Keep old release
binaries and populated workspaces untouched.
