> Historical implementation reference. Current procedures are in [the user guide](../docs/Swift_NAVFIT_User_Guide.md); [current state](CURRENT_STATE.md) identifies the verified build. Standalone comments/review windows and resolved/reopened note threads described below are superseded by independent report notes/replies and embedded report tools.

# Report review comments

> Current native navigation: use the [workspace guide](UX_WORKSPACE_GUIDE.md) and
> [current state](CURRENT_STATE.md). This guide retains service details and
> earlier verification history; legacy screen instructions may differ.

Swift NAVFIT 0.7.0 adds report review threads, replies and resolution history. Open a saved draft, scroll below its narrative section and choose **Review comments on saved draft**.

Write a comment and choose **Post new thread**. The thread references the latest saved revision at posting time, excluding pending editor values. Select a thread to read its full text and replies, then use **Post reply**, **Resolve thread** or **Reopen thread**. Replies retain the original thread's report and revision reference even after newer report revisions exist. Resolve/reopen actions append attributed events; repeating the same action does not create another event. Resolved threads must be reopened before adding replies.

Comments and events show the active local contributor and workstation-local time. Profiles are attribution, not authentication. This release adds no profile switcher; the test fixture uses two local profiles to verify attribution. Threads list their saved revision number, author and date. The selected thread displays its saved revision ID and full text, followed by resolution history.

Review comments are stored separately from OfficialComments and do not change report snapshots, official narratives or report workflow state. Posted comments have no edit/delete action. Return to open editor preserves unsaved editor text. Archived reports allow reading comment history but reject new comments, replies and resolution changes until recovered.

Comments are revision-level review notes. Field-level anchors, reviewer assignments, due dates, routing, notifications and approval tracking remain pending. Comment entry has no automatic recovery checkpoint; post it to retain it. No export or policy-validation capability is added.

The blank distribution uses schema 7. Carry an existing workspace into it with `scripts/upgrade-workspace.ps1 -SourcePath <existing workspace> -DestinationPath <new workspace>`. The script defaults to the 0.7.0 template, verifies copied rows and fields, preserves the source file and refuses destination overwrites. Do not replace a populated workspace with the blank distribution.

Verification: 28 focused comment assertions including long text, multiple contributors, report/revision isolation, blank rejection, replies, resolve/reopen history, archived guards, no official revision/narrative writes, actual form callbacks, pending editor preservation and reopen. Also 54 workflow, 24 provenance and 46 smoke assertions pass. Schema-6-to-7 and schema-7-to-7 migrations compare every original row and field and reopen. The delivered master has 18 empty user-data tables. Full visual inspection and Access Runtime/x86 remain untested.
