# r99 — pending comparison cell recovery

Highest-priority item from REVIEW_R98_2026-10-10.md implemented. Application 0.25.0-dev/schema 25; no schema change. r98 remains the immutable published baseline. These are separate local candidates.

Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-10-r99.accdb` for synthetic practice, the matching blank for a new workspace, or `Swift_NAVFIT_UI_TestWorkspace_2026-10-10-r99.accdb` to retain the current r97 workspace data.

## Changes

- Queued cells are removed only after successful commit. A failed cell and all later cells retain their intended values and original contributor; earlier saved revisions stay saved.
- Failure stops the automatic timer. The comparison displays the remaining unsaved-cell count and offers **Retry pending cells** and **Discard pending cells**. Correcting a cell coalesces its queued value; retries do not replay earlier successful commits.
- Native Form_Unload cancels closure when draining fails. The bound cache and queue remain available. Group-page navigation respects that cancellation.
- Contributor switching is blocked while cells are pending. Each queued cell also carries its original contributor, which is verified before saving.
- Refresh, proposal staging/application and report creation cannot bypass pending cells. Explicit discard restores the saved comparison cache and requeries native controls; already saved revisions remain intact.
- Focus moves from a recovery action before that action is hidden, avoiding the Access hide-focused-control error.
- In-app version/schema Help uses APP_VERSION and SCHEMA_VERSION instead of stale literal values.

Normal valid edits still save automatically, and normal closure remains prompt-free. Local approval reopening retains its existing explicit choice. Choosing not to reopen leaves the cell available for retry or explicit discard.

## Verification

42 native assertions cover a multi-report queue with a successful first edit, invalid second edit and later edit; partial-save preservation; correction/coalescing; retry without duplicate revisions; native close cancellation/cache preservation; group-page cancellation; proposal guard; contributor guard and attribution; injected DAO field-validation failure/rollback; explicit discard while its button has focus; successful final close and cache deletion.

Existing queued-cell, 26 approval/rollback/retry, all-three-type control-event, blank/demo startup and 354 author-property/navigation checks pass. Data comparison verifies all 49 user tables unchanged between r97 and the r99 preserved workspace. Current additional group-journey results and hashes are recorded in PENDING_CELL_RECOVERY_R99_VERIFICATION.json.

No computer use. Pending comparison recovery is retained within the open Access session; this does not add startup recovery of queued comparison cells after a process or computer crash. Existing report-editor recovery is unchanged. Physical keyboard/screen-reader/DPI acceptance remains outstanding.

## Next priorities

Live Members/snapshot lifecycle optimization, then change-driven validation, saved aggregates and checkpoint writes. No snapshot optimization or returned-group/narrative-fit implementation is claimed in r99.
