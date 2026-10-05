# Focused review and Reports return (r23)

Report editor > Review > Review routes opens one saved-report review workspace.
Group > Review > Review selected member opens the same view with group context.
Route labels show template, saved revision number, status and date. The step
list names contributors; selecting a step shows its full instructions and
previous decision. The header always uses saved report values, including when
pending edits remain in the editor. Review actions retain report scope.

The screen explains the current step, assigned reviewer, stale state, local
approval lock and next action. Disabled-action guidance states contributor and
workflow requirements. Route templates remain in Resources; no capability was
removed from the existing administrative workbench.

Restart stale route creates fresh pending steps using the original template
and current saved report/group baseline. Only its requester can restart a stale
pending or accepted route. The old route and decisions remain as cancelled
history. A changed displayed baseline, archived group/template, removed member
or failed replacement blocks the operation. Cancellation and replacement share
one DAO transaction, including the pending-route count query; Access domain
aggregates cannot see an uncommitted cancellation.

Opening a report from Reports captures selection and native vertical scroll
position with the top visible report identity. Existing filters, search and
view remain in the open list. On return, the list refreshes, restores selection
if still present, and anchors scrolling to that report after save reordering.
If the anchor disappears, it uses a bounded previous position. If selection
falls outside filters, it clears selection instead of selecting another report.
An Access timer restores after native close/focus processing; the state is
session-local and adds no schema fields.

Automated scratch checks exercise actual OGrid scrollbar messages, a 90-report
fixture, native editor close/autosave, route restart, rollback, contributor
restrictions, group staleness, approval and reopening. Manual visual, keyboard,
screen-reader and display-scaling acceptance remains outstanding.

Final focused suite: 133 assertions. Existing workflow suite: 49 assertions plus two forced-shutdown cases. Major workflow suite: 44 assertions. All 49 source standard modules match both final native artifacts. Four repeated authored viewport widths are checked for horizontal bounds; those checks do not establish physical Windows scaling acceptance.
