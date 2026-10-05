# Shared blocks and editable comparisons - 2026-10-04

This implements review sequence step 3 in native Access. Version remains
0.22.0-ux-preview and the workspace schema remains 21.

## Shared report blocks and defaults

- Group bulk edits and group defaults use the report editor's official block
  order, form-specific captions, width rules and adjacent checkbox layout.
- Four contextual sections replace eight-fields-per-page entry. Fields live
  in a scrolling native subform, with text, date, choice and checkbox controls.
- Shared saved, blank, mixed, staged and explicit-clear states have text labels.
  Each field has a Clear button, including fields that initially have no value.
- Only edited values are staged. Choose a section, edit, enter a reason, Preview,
  inspect complete before/after values, then Apply. Changing inputs or reason
  invalidates Apply. Group membership and saved-report changes require reopening.
- Group defaults follow the same journey. Applying them changes suggestions only;
  reports retain their revisions. The preview captures full before/after values,
  retains them for Results and rejects concurrent default-version changes.
- Default writes use one outer transaction around the existing audited default
  service. Bulk writes still use existing atomic report commits and guarded undo.
- Command, folder and report defaults remain available through existing tools.

## Editable trait comparison

- A continuous native subform shows member, local row state, seven editable
  trait cells, ITA and promotion recommendation. It contains no comments.
- Column captions follow each form's official trait blocks. Hover help carries
  the complete caption. Columns adapt to the available panel width; comparison
  display order follows the saved meeting order.
- Cells stage durable proposals. Revisiting a cell supersedes its previous
  pending proposal; it does not leave duplicate active proposals. Returning to
  the baseline removes the active proposal. Superseded entries remain auditable.
- Pending ITA and group trait average are arithmetic previews. Numeric traits
  contribute; blank and NOB marks and whole-report NOB reports are excluded.
- Archived, locally locked, changed-report and changed-group baselines reject
  edits with a next action. Preview shows member, official block, complete before
  and after values and state. Apply retains the existing atomic revision guards.
- Results retain applied before/after values and an explicit success message.
  Apply is disabled for the historical comparison. Return to group closes it.
- After Apply, the comparison is historical. Reopen a group comparison to work
  from its new saved values. No manual roster-freeze step is required.
- Editable rows use a uniquely named per-session temporary Access database.
  Saved reports are never bound to this grid. Native unload detaches the rows
  and removes the temporary cache; staged proposals remain in the workspace.

## Verification

The focused native harness passes 90 assertions covering all three report types, mixed values, clear
staging, defaults without report commits, concurrent defaults, direct multirow
edits, proposal replacement, ordered reload, full previews without parameter
requests, pending averages with whole-report NOB exclusion, archive/approval
locks, stale batch rollback, atomic Apply and native cache cleanup.

The test runner is scripts/test-ux-block-tools.ps1 and is required by the
seeded preview builder. Exact final artifact results are recorded in
planning/CURRENT_STATE.md. Existing editor, workflow, responsive, senior,
major-feature, shell and populated-demo checks are also run on scratch copies.

Native visual review at actual Windows scaling and keyboard/screen-reader
acceptance remain separate from these behavioral and geometry checks.

The final audit also corrected zero-argument calls before a colon: VBA can
parse a bare function name there as a line label. Explicit Call statements
ensure initial averages, viewport refresh and post-Apply row reload actually
execute. The harness verifies initial ITA and group average on opening.
