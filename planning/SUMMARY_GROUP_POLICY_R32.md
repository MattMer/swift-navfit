# Summary-group policy and consequence previews - r32

Development build, version 0.23.0-preview, schema 21. This does not replace the
published prerelease. Open `output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r32.accdb`
for synthetic reports, or `build/ux-r32-check-blank.accdb` for an empty workspace.

## Implemented

- Group workspace **Policy review** records enlisted UIC grouping choices or
  officer normal/dual-hatted grouping, manually confirmed officer Active/Reserve
  category basis, a required review note, contributor and timestamp.
- Competitive categories now cover the instruction's Active/Reserve LDO and CWO
  splits, including the 740x and 7371 exceptions. Block 5 does not automatically
  establish an officer's competitive-category basis.
- Combined enlisted UICs require an explicit recorded choice. Separate enlisted
  and dual-hatted officer grouping check member UIC against reporting senior UIC.
  The application cannot establish flag-officer/SNDL authority; the reviewer
  records the applicable basis.
- Completeness is deliberately unchecked on opening the review. Recording it
  does not waive findings or certify Navy compliance. Report revisions,
  membership, workflow state and policy changes invalidate the review baseline.
  Pending member edits also prevent recording completeness.
- Membership changes require **Preview**, then **Apply previewed membership**.
  Selection, notes, candidate report revisions and group changes invalidate the
  preview. Existing atomic membership services remain intact.
- Membership and bulk previews show before/after report and observed counts,
  graded-trait arithmetic average, promotion-capacity findings and the need for
  a fresh roster review after applying changes. Empty observed rosters have no
  capacity calculation. These previews do not save report revisions.
- Policy context uses versioned XML in existing `SummaryGroup.KeySpec`; unknown
  legacy content is rejected without overwrite. Changes are journaled with full
  before/after values. No schema change or roster snapshot is required.

## Evidence and limits

Policy locators and excerpts use the existing fingerprinted BUPERSINST 1610.10H
CH-2 evidence and its verification script. Category patterns do not establish
that a designator exists or that a member's commission status is accurate.
The user remains responsible for the complete eligible command roster.

The targeted native Access harness passes 39 assertions covering NOB exclusion,
unresolved/mixed keys, category splits, UIC context, stale policy review,
read-only consequences and membership preview guards. Additional regression
results are recorded in CURRENT_STATE.md. COM form checks establish functional
behavior, not physical-display, keyboard or screen-reader acceptance.

## Still outstanding

- Exercise real leadership review scenarios and independently verify boundary
  capacity/category examples against the instruction.
- Manually verify new forms on target displays and with keyboard/screen reader.
- Verify group-policy data preservation through supported workspace-transfer
  journeys before publishing a new deployment.
- Command-wide roster completeness remains a human attestation; the app has no
  external personnel source from which to prove completeness.
