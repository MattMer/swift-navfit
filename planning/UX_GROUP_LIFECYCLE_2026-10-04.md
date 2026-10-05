# Summary-group lifecycle - 2026-10-04

Native summary-group rename, archive and recovery implement the next lifecycle
recommendation. Version remains 0.22.0-ux-preview; schema remains 21.

- The Summary groups list has Active, Archived and All filters, report counts,
  explicit status, and contextual Rename, Archive and Recover actions.
- Rename uses a focused dialog with a 120-character title limit. Failed or
  stale saves keep the pending name available for correction.
- Archive is reversible and hides the group from the active list. It retains
  members, defaults, snapshots, comparison proposals and operation history.
  Member reports retain their own archive/review state and saved revisions.
- Archived group workspaces permit inspection and opening reports; group
  management, defaults and comparison actions explain recovery requirements.
- Recover restores group management without recreating members or report data.
- Each successful lifecycle change increments the group version and records an
  audited operation in one transaction. Older bulk and comparison baselines
  reject Apply after a change, including a rename or archive/recover cycle.
- Selection is cleared when its group leaves the current filter. Keyboard
  selection updates action availability. Actions move focus safely before
  disabling the control that initiated the change.

The focused native suite passes 41 assertions for dialog save, validation,
filters/status/counts, retained data, workspace locks/recovery, stale writes,
stale bulk/comparison previews, audit entries and preservation of failed input.
The seeded demo builder now requires this suite before publishing a review file.
See CURRENT_STATE.md for the final demo and regression results.

Transparent calculation explanations, checklist navigation and remaining
review/disabled-action wording are subsequent recommendations. Physical DPI,
keyboard and screen-reader acceptance remains separate from automated checks.
