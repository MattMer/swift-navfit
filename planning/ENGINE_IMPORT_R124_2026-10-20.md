# Database-engine import verification - r124, 20 October 2026

The prior release bulk-copied tables but then retrieved and compared every field of every record in VBA. On a provenance-heavy workspace that remained the dominant cost. The new normal path uses database-engine queries for exact comparison. It does not substitute sampling or row counts for value verification.

## Implementation

- Source queries are prepared outside the data transaction, read through read-only DAO, and shared with relationship restoration. Source forms/code are never opened. Temporary queries are removed after success or failure.
- Each table requires matching counts and a primary-key join (including composite keys). Indexed numeric-key batches select at most 8,192 first-key positions, including any ties. Up to 12 field predicates are combined per query to avoid Access expression-complexity limits on wide report tables. Each source row belongs to one numeric-key range; each field is covered by one field partition. Small nonnumeric-key tables use one row range.
- Queries return only counts or the first mismatch; VBA no longer walks all record values. Binary StrComp compares complete text/memo values, including case, Unicode and trailing spaces. Null/empty and Null/zero remain distinct. Numeric, date, currency and boolean values compare directly. Missing/extra rows and equal-count key substitutions are rejected. Unsupported types or missing primary keys fail closed.
- The importer becomes modal while busy and disables all its command buttons. Apply, source preview and source chooser have reentry guards; the native unload guard blocks closing during import. Guarded DoEvents calls occur between completed queries/batches so Windows can process paint/focus events. Controls and original modality are restored on success or failure.
- This is cooperative event processing, not a background worker. A single ACE copy/query or file-fingerprint call still runs synchronously. The bounded verification batches address the long unresponsive verification loop; no claim is made that every storage operation is instantaneous.

## Realistic performance evidence

A separate synthetic workspace contains 262,144 ImportRawValue records with text, Unicode, quotes, newlines, trailing spaces, and Null values. The source is approximately 180 MB. The exact dated r124 Blank imported it through the real Settings & Help Apply path in **12.804 seconds**: bulk copying 3.52 seconds, full field verification 8.73 seconds, plus fingerprint/commit work. Preview/startup/template-copy time is excluded. A native timer ran 46 event turns during import and verified that duplicate Apply/preview/chooser calls and busy-form closure were blocked.

The previous r123a importer was tested on the same synthetic source. It was still running beyond nine minutes and its owned disposable comparison process was intentionally stopped. No completed old-runtime measurement or exact speedup ratio is claimed. The aborted comparison's RPC error is from terminating that scratch process, not from the new importer. No user database session was terminated. Final source fingerprints are unchanged. These are local synthetic timings, not a guarantee for another workstation or storage/synchronization environment.

The older small fixture had not represented this workload. It remains useful for regressions: 20,856 revision values, long Unicode notes, out-of-order parents, next AutoNumber, late relationship failure cleanup/retry, and oversized-memo rollback passed. Engine queries add some fixed overhead on tiny datasets; the primary goal is removing per-value VBA work and excessive progress updates at real provenance volumes.

## Exactness and release checks

Sixteen adversarial native cases passed: full 70,000-plus-character memo equality; case mismatch beyond 70,000 characters; memo/short-text trailing spaces; text case; Null/empty and Null/zero; timestamps; precise currency and floating values; booleans; Unicode accents; composite-key substitution; missing/extra rows; restored equality. A separate timer injected a target report-field change inside the real migration transaction. Engine verification rejected it before commit, rolled back imported data and setup, restored controls/modality, and allowed a valid retry.

All 26 retained schema fixtures passed through production migration, including the current-schema UI. Source fingerprints, bootstrap rollback, future-schema/unknown-table/changed-source/populated-target guards passed. Populated schema 20 and 26 upgrades preserve original fields and reopen; canonical replies/private drafts and out-of-order parent IDs survive. Native 48 smoke and 39 Demo checks pass. The exact Blank has no user-data rows; Demo retains 13 synthetic reports and 5 groups. Original populated r116 fingerprint remains 6ACE66A9C2883E1BE6CBCFBEC584249B254A61620D998486468E205BCECBEFE2.

## Delivery and repeatability

Files are `output/Swift_NAVFIT_Blank_2026-10-20-r124.accdb` and `output/Swift_NAVFIT_Demo_2026-10-20-r124.accdb`. Both are uploaded to the designated Drive folder; remote names, sizes and parent IDs were read back. Access is closed. Physical display/DPI/accessibility/printing acceptance remains separate and deferred.

- `scripts/test-engine-import.ps1 -BuildPath <blank> -SkipBaseline` generates the 262,144-row synthetic fixture and runs native exactness/UI checks. `-SourcePath <synthetic source>` reuses a fixture. Omit SkipBaseline only when deliberately allowing the old comparison to run for a potentially long time.
- `scripts/test-engine-verification.ps1 -BuildPath <blank>` runs the 16 mismatch cases. Add `-SourcePath <large synthetic workspace>` for injected-corruption rollback and retry.
- Existing `test-swift-data-import.ps1`, `test-workspace-upgrade.ps1`, `test-report-message-import.ps1`, and `test-bulk-workspace-import.ps1` retain broader compatibility and rollback coverage.

See ENGINE_IMPORT_R124_VERIFICATION.json and output/CURRENT_FILES.json for hashes, test evidence and delivery links.
