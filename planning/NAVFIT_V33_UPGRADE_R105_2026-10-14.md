# Automatic NAVFIT v33 → Patch 2 import conversion — r105

14 October 2026. Application 0.25.0-dev; Swift schema remains 25. Local candidates only; GitHub r104 remains the published prerelease.

## Try it

- `output/ux/Swift_NAVFIT_UI_Demo_2026-10-14-r105.accdb` — existing synthetic reports/groups plus the latest import behavior.
- `output/ux/Swift_NAVFIT_UI_Blank_2026-10-14-r105.accdb` — new workspace.

Use Resources → Import NAVFIT database as before. A recognized older v33 source is converted automatically before the source-group chooser/preview. No separate converter or confirmation is required. Review and Apply still control the actual import. Single-report review/import uses the same preparation path.

## Changes reproduced from the supplied utility

Inspected the supplied `NAVFIT98A_v33p2_Upgrade_Tool.accdb` form VBA on a separate copy, with automatic code execution disabled. Swift does not open or execute that utility during imports. [Recorded tool specification and checksum](NAVFIT_V33_UPGRADE_TOOL_EVIDENCE.json).

| Table | Added fields and text lengths |
|---|---|
| Reports | DoDID (10), RSDoDID (10) |
| Folders | RSAddress1 (30), RSAddress2 (30), RSCity (15), RSState (2), RSZipCd (9), RSPhone (10), RSDSN (7) |

Missing fields are added. NULL values in all nine patch fields become empty strings; existing non-NULL values remain. This adds native compatibility columns only; Swift's member editor still uses SSN.

Detection validates every required native field's name, type and size, allowing only the nine known omissions (including partial upgrades). Missing unrelated fields, unexpected columns, incompatible existing patch fields, and linked/hidden source tables are rejected. This is support for recognized v33 schemas, not arbitrary pre-v33 databases. Raw staging files and export templates still require the full supported schema.

## Original files, provenance and exports

The original database is opened read-only. Only a GUID-named file in the temporary directory is converted, in an isolated DAO workspace and transaction. Already-current files use the normal read-only path. Read-only original files are supported. The temporary conversion is removed when the review closes or is replaced; failed conversions are removed and cannot partially import records. A process crash may leave a temporary copy for later operating-system temp cleanup.

Original path and original-byte hash remain in import receipts and duplicate checks. Source changes and prepared-copy changes are checked before preview/apply. Operation history records automatic conversion, explaining why the retained Patch 2 raw row includes newly initialized compatibility fields not present in the original. No Swift schema migration is needed.

The utility appends fields, whereas Swift's earlier validator/rebuilder required their factory ordinal positions. Validation/rebuild now matches exact known names and types, retaining source ordinals in provenance and checking contiguous ordinals, unique field names and complete rows. This allows upgraded imports to export using a standard Patch 2 template without dropping native-only values. The original database is never overwritten or replaced by an upgraded sibling.

## Verification

- Exact nine field definitions and nine NULL-initialization statements matched against the supplied utility's exported VBA.
- 134 new assertions: current/legacy/partial/read-only sources; invalid type/size, missing unrelated field, unexpected field and failed initialization rejection; original preservation; copy cleanup; group import and export; original receipt identity; duplicate protection; source/prepared-copy stale guards.
- 507 existing adapter assertions; 410 exchange assertions; 99 stale/hash/identity/lock/undo guards; 247 selected-group import/export/UI assertions. Total: 1,397 assertions.
- Both candidates compiled. Normal blank startup opens setup; demo startup opens Overview. Blank user tables checked empty. [Exact artifacts](NAVFIT_V33_UPGRADE_R105_VERIFICATION.json).

Tests use synthetic older schemas made by removing the documented patch fields from a supported synthetic database, including partial and malformed variants. No original historical v33 sample or NAVFIT98A GUI round-trip was supplied/tested in this change. Printing and returned-group reconciliation are unchanged. Existing single-report returned-file tests pass; older returned-file reconciliation is not a newly added migration workflow.
