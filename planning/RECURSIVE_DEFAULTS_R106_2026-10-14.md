# Direct recursive defaults — r106 (14 October 2026)

Local candidate: app 0.25.0-dev, schema 25. Includes r105's automatic older NAVFIT import conversion. GitHub's published release remains r104; this work has not been pushed.

## Behavior

- Global and summary-group defaults use the report-shaped field editor. Inheritance/status explanations and per-field Reset/Clear buttons have been removed. To blank a value, edit its field normally.
- Finishing a changed field asks once whether to apply it recursively. Confirmation names the field, new value, scope, affected report count and any local approvals returning to In-Work. Existing customized values are replaced for that field.
- Global changes save the command default and all active group defaults, and update applicable fields in their active reports. Group changes save that group's default and update its active members. Future groups/reports inherit the saved defaults. Global scope excludes ungrouped reports; archived groups/reports are excluded.
- Cancel restores the previous field value. Validation or storage failure restores the control and rolls back the entire operation, including defaults, saved revisions and approval reopening. Reports with the same value get no extra revision or approval reset.
- There is no defaults Preview/Apply stage or pending batch that blocks navigation. Preview controls and inheritance/reset/clear controls are physically deleted from generated forms. The shared-field host and hidden value buffers remain because adapters use them.
- Reporting-senior directory application now confirms once and applies related fields together, with one revision per affected report. Its backend transaction/batch history remains; there is no preview screen requirement.
- Import previews and comparison proposal review are separate workflows and retain their safeguards.

## Candidates

- Synthetic demo: `output/ux/Swift_NAVFIT_UI_Demo_2026-10-14-r106.accdb`
- Blank: `output/ux/Swift_NAVFIT_UI_Blank_2026-10-14-r106.accdb`

Existing databases were not overwritten. These are separate local candidates, not an update to a workspace already open on another filename.

## Verification

387 native Access assertions passed, plus recovery-timer and forced-shutdown scenarios. Native Access scratch-copy suites cover recursive scope, EVAL/CHIEFEVAL/FITREP, date/checkbox normalization, cancellation, future inheritance, stale plans, approved-report reopening and history retention. An injected failure on a later report proves earlier revisions/defaults/approval changes roll back; retry succeeds. Navigation/save/recovery and reporting-senior suites also run against the candidate. See `RECURSIVE_DEFAULTS_R106_VERIFICATION.json` for exact final hashes and assertion counts.

The screenshot's syntax error was an unused test variable named `shared`, a reserved VBA word. It was removed; the harness and candidates compile. Joined report ordering was qualified to avoid ambiguous `ReportID`. Older test harnesses were updated for immediate changes and the direct member-list navigation introduced earlier.

Automated tests inject Yes/No into the same confirmation handler. They do not establish manual modal-dialog, physical display/DPI, keyboard or screen-reader acceptance. Printed-fit and NAVFIT handoff limitations remain unchanged.
