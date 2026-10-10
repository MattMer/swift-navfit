# RUAD import and refresh — r97, 9 October 2026

Open output/ux/Swift_NAVFIT_UI_TestWorkspace_2026-10-09-r97.accdb. This copies the confirmed r96 workspace without seeding or rewriting its records. Existing workspaces remain preserved; schema stays 25.

## User journey

Resources → Command member directory → Import / refresh RUAD CSV.

Choose the supplied 35-column CSV. All rows initially Skip. Select new profiles for unmatched names is an explicit bulk selection; existing names remain skipped until matched. Select a row to choose Create new, Refresh existing with an explicit directory member, or Skip. Vacant billets cannot be imported as people. Names and billet identifiers do not establish identity.

Preview selected changes. Selecting a row then shows all 35 full before, incoming and after values, including preserved conflicts. Changed selections disable Apply and invalidate the prepared service plan. Apply reviewed changes commits all selected profiles atomically and reports the results. The operation is recorded in Settings & Help → operation history. Closing before Apply writes nothing.

New profiles use actual enlisted DESG or officer RATE/RANK as their suggested report rank/rate; no numeric conversion of identifiers. SSN remains unentered. Names too long for the official-name capacity leave that separate field unentered; no truncation. Raw officer designator is retained, not automatically transferred to report blocks. CSV source text is UTF-8 and quoted commas/newlines/escaped quotes are supported.

## Refresh safety

Refresh requires explicit matching. It preserves display name, report name, report rank/rate, SSN, report links and prior evaluations. Roster fields use the last imported source values as a three-way baseline: manual changes are preserved unless Use incoming RUAD values is explicitly selected. All resulting values appear in preview. A nullable-source baseline and source path/time are retained within existing RUADData XML; manual profile saves retain that provenance. No additional schema or table is needed.

Duplicate/repeated selections, changed directory versions, changed files, unsupported/malformed CSVs and late record failures reject the operation. No absent members are deleted or archived. Same-name people require explicit matching to separately established directory profiles; names alone cannot create duplicate profiles in this batch workflow. Automatic batch undo and missing-member reconciliation are not implemented; existing import undo services are unchanged.

## Verification

- Native compile.
- 71 sample mapping, provisional identity, all-column preservation, three-way/manual override, duplicate, stale-source/version, late-failure rollback and native selection/preview assertions.
- 26 approval, rollback, retry, repeated trait-save and reopen assertions.
- All 49 user tables equal r96 before user import.
- 354 authored-property, obsolete-control absence, normal startup and direct populated-group navigation checks.

Sample: 66 rows, 60 people, six vacancies. Tests import only into scratch copies. Exact interactive trait-save fix was confirmed by the user in r96. r97 retains that fix and local diagnostics. No computer use, GitHub publication or release claimed. Physical keyboard/screen-reader/DPI acceptance remains separate.

Also corrected the summary-group archive filter caption from Group status to Archive status; the workflow-status tool is separate. New import controls use the established Siltstone palette and Segoe UI. Full/incremental admin builders now install the new modules/forms through apply-ux-ruad-import.ps1.
