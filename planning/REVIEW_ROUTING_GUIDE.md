> Historical implementation reference. Current procedures are in [the user guide](../docs/Swift_NAVFIT_User_Guide.md); [current state](CURRENT_STATE.md) identifies the verified build. Standalone comments/review windows and resolved/reopened note threads described below are superseded by independent report notes/replies and embedded report tools.

# Local review assignments and My Work

> Current native navigation: use the [workspace guide](UX_WORKSPACE_GUIDE.md) and
> [current state](CURRENT_STATE.md). This guide retains service details and
> earlier verification history; legacy screen instructions may differ.

Swift NAVFIT 0.8.0 adds a first local routing step. Review assignments reference a saved report revision, requester, reviewer, optional due date and instructions. Report editing stays available. Review outcomes do not change report workflow state or constitute official approval, signature or export readiness.

From Home, open **My Work and local contributors**. Add a named reviewer if needed. Contributor names must be nonblank, at most 120 characters, and distinct from existing names. Select an existing contributor and Switch contributor to use that person's pending-work list and record activity in their name. Close the draft editor and review comments before switching. Local profiles provide attribution, not authentication. Use one person at a time in the same authoritative workspace file.

To assign a report, open its saved draft and scroll below the narrative and comments actions to **Assign saved draft for review**. The assignment screen identifies the report and saved revision ID. Select the reviewer, optionally enter a due date and add instructions, then choose **Assign saved revision**. Pending editor values are excluded. Refresh updates the saved-revision context; the commit checks that it has not changed again. One pending review per report is allowed. A requester can cancel it before making another assignment.

My Work defaults to pending assignments for the active contributor. **Show all review history** also shows finished tasks and other contributors' tasks. Rows show title, reviewer, state, due date, Current/Stale and archive flag. Select a task for requester, due date, overdue indication, revision ID, instructions and attributed outcome history. **Open selected draft** requires the current editor to be closed; it opens current report values. To inspect differences from the assigned revision, use the editor's existing revision comparison.

Only the assigned reviewer profile can Accept or Return a pending review. A return requires a reason. Only the requesting profile can Cancel it. Accepted means this local review step was accepted; it is not a Navy approval. All outcomes append an event and preserve the original assignment. Completed tasks cannot be overwritten; create another assignment for another review pass.

**Since 0.52 (F01):** a pending review's assignee holds the report and may edit and save it before accepting or returning; other contributors' saves are refused until the review finishes or is cancelled. Finishing records the revision (ReviewTask.FinishedRevisionID). Ordered routes follow the same model with Forward, Send back and Pull back; see the user guide, chapter 12. Earlier text about stale reviews describes behavior before 0.52. Archived reports reject assignment and reviewer outcomes; the requester can still cancel pending work. Comment-only activity does not change the saved revision and therefore does not make an assignment stale.

Notes in the assignment/outcome textbox have no automatic checkpoint; submit them to retain them. Due dates are local calendar dates; audit timestamps use workstation-local time. Sequential multi-step route templates, assignment reassignment, notifications, reviewer authorization, official approval locks and export readiness remain pending. Exclusive use is advised but not enforced in this prototype.

The blank distribution uses schema 8. Run `scripts/upgrade-workspace.ps1 -SourcePath <existing workspace> -DestinationPath <new workspace>` to carry existing data into the new app. The default template is the 0.8.0 delivery. The source is preserved and destinations are never overwritten. Do not replace a populated workspace with the blank distribution.

Verification: 38 focused routing assertions, including long instructions, contributor switching, wrong-reviewer rejection, duplicate pending work, missing return reasons, stale/archived guards, form callbacks, pending editor preservation and late injected assignment/outcome failures with full rollback. Also 54 workflow, 28 comments, 24 provenance and 46 smoke assertions pass. Schema-7-to-8 and schema-8-to-8 migrations compare every original row and field, preserve source hashes and reopen. All 20 user-data tables in the delivered master are empty. Full visual inspection and Access Runtime/x86 remain untested.

Testing corrected cleanup that attempted to close an already-closed recordset. Routing handlers now roll back first and track closed recordsets; transaction reads and event writes use one retained DAO database handle. Failed candidates were not published.
