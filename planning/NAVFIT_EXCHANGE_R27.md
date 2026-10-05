# NAVFIT exchange r27 — 5 October 2026

## Implemented scope

The native Access preview imports all supported Swift report values for EVAL,
CHIEFEVAL and FITREP, exports one saved report, and reconciles its returned NAVFIT
file against a durable export baseline. Version remains `0.22.0-ux-preview`,
schema 21. Operational workspaces and the supplied test source remain unchanged.

Each form maps 74 saved values plus report type. Swift's other three saved keys
are local title, display name and draft notes. All **128 Reports columns**, 78
Folders columns and six Summary columns are retained as typed, null-aware
provenance. Native-only content is preserved rather than presented as an editable
Swift field. The catalogue is `src/navfit-field-map.json`; bindings and codes
were checked against the supplied schema, executable and dropdown database.

| Form | Native columns for blocks 33–39 |
|---|---|
| EVAL | PROF, QUAL, EO, MIL, PA, TEAM, LEAD |
| CHIEFEVAL | LEAD, TAC, PROF, MIS, EO, TEAM, MIL |
| FITREP | PROF, EO, MIL, TEAM, MIS, LEAD, TAC |

Trait code 0 maps to NOB; 1–5 map to 1.0–5.0. Promotion codes 0–5 map to NOB,
Significant Problems, Progressing, Promotable, Must Promote and Early Promote.
Typed dates are preserved. A `yyMMMdd` counseling date requires an explicit
1900/2000 century. Unknown encodings/codes, unknown checkbox values and oversized
native fields fail visibly; values are never guessed or truncated. Structured
name parts and independently stored full names remain independent.
Date fields containing a time are rejected rather than silently losing it;
counseling export supports explicitly resolved years 1900–2099.
Every export resets native `IsValidated` to False. A Swift draft or returned
native validation flag does not certify the newly exported file.

## Journeys

Reports → **Import reports** opens Choose → Preview → Apply → Results for full
supported NAVFIT report import. The chooser also opens **NAVFIT export and
returned-file reconciliation**. Select a saved report, the verified
`NAVFIT_BLANK.accdb` template and a new export filename; close the editor and
resolve recovery first. Export creates a native `.accdb` and a companion
`.accdb.swift-baseline.accdb`. Send only the native file and keep the baseline
locally. Existing filenames are never overwritten. Failed/pending operations
are recorded and never reported as completed exports.

Export reconstructs preserved native content without requiring the original
source file. The companion freezes Swift and native values and is sealed by the
workspace operation journal. Return preview shows all report columns with full
baseline/current/returned values. Unchanged returned values preserve later Swift
edits; distinct edits on both sides require **Keep current** or **Accept returned**.
Resolved rows update immediately. Apply also accepts the displayed native-only
source changes; cancel if those are unwanted. Normal atomic revisions, source
receipts, provenance and compensating undo remain the write path.

Identity uses originating workspace/report GUIDs and exported native ID/type,
never an SSN guess. Modified baselines, changed files after preview, stale Swift
revisions, unresolved conflicts, approval locks and superseded native receipts
stop Apply. Added/deleted/renumbered returned reports require a new import.

## Verification

Installed and supplied NAVFIT executables both report **2.2.0.332**, SHA256
`5366F1C1BB147F1706094C5E3E1D62981F34FC6CB301E4006A892ACFBFAB58CF`.
Native UI checks used the installed application: exported reports opened, front
and back pages were inspected, and comments were edited and saved in each form.

| Check | Result |
|---|---|
| Mapped values and assessment conversions | 497 assertions passed |
| Single-report export, all 128 columns, simulated conflicts | 410 assertions passed |
| Stale/hash/identity/lock/undo/missing-source/new-draft/date guards | 98 assertions passed |
| Native import/return UI, including resolved-row feedback | 28 assertions passed |
| Actual EVAL return and re-export | 75 decoded values + 128 native columns matched |
| Actual CHIEFEVAL return and re-export | 75 decoded values + 128 native columns matched |
| Actual FITREP return and re-export | 75 decoded values + 128 native columns matched |

Native proof files are under
`build/tests/navfit-adapter-e65564e9-13fc-4f8b-bcaf-e114d3de087c`:
`export-2.accdb` (CHIEF), `export-3.accdb` (EVAL), and `fitrep-native.accdb`.
Each `.native-proof.json` names and hashes its verified re-export. CHIEF exposed
an empty-text/null reconstruction bug; it was fixed and the re-export then
matched. Automated fixtures also prove long Unicode Memo preservation.

NAVFIT Save itself normalizes some values. An earlier three-report observation
compared 824 cells and recorded 41 changes, including summary counts, empty trait
descriptions, masked contact cleanup and CHIEF legacy-name parsing. Return
preview exposes these changes and provenance retains them. This is not a claim
that NAVFIT leaves arbitrary synthetic content unchanged or passes Navy checks.

`build-preview.ps1` now gates demos on NAVFIT services/UI alongside existing PDF,
validation, review, My Work, groups, responsive editor, recovery and startup
suites. Generated evidence is under `planning/evidence/navfit` and `build/tests`.
The checked-in [verification manifest](NAVFIT_R27_VERIFICATION.json) records the
final artifact hashes and the three actual native return/re-export receipts.

A clean `build.ps1` rebuild also passes installed-source and native event-handler
verification (57 modules). During verification, Access dropped newly inserted
form class code when the form was saved before closing. The builder now binds
unload events after construction and closes those forms with `acSaveYes` to
persist the handlers. Intermediate construction no longer attaches empty event
procedures.

## Remaining limits

- Export is one report at a time. Complete group/batch export and group promotion
  counts/RSCA reconstruction remain separate work.
- Native-only columns are retained and reviewable, not new editor fields.
  Legacy non-date counseling text is rejected.
- Changed identities, added/deleted reports and returns after newer native
  receipts are not guessed; obtain a fresh export baseline or import separately.
- Physical keyboard/screen-reader/scaling acceptance for the exchange screen,
  Access Runtime/x86 and other NAVFIT versions remain unverified.
- Printed layout, official signatures and complete Navy export-readiness
  certification remain outside this milestone.

Use the newest artifacts in `CURRENT_STATE.md`; earlier r27 candidates remain
available for evidence but precede the final preservation fixes.
