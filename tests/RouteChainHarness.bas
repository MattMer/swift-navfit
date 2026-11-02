Attribute VB_Name = "RouteChainHarness"
Option Explicit
'F01: a routed report is one shared draft held by the current step's reviewer.
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 1860, , message
    checks = checks + 1
End Sub
Private Sub Edit(ByVal id As Long, ByVal key As String, ByVal value As Variant)
    Dim changes As Object
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add key, value
    Call CommitSavedFields(id, CurrentRevision(id), changes, "Route chain test")
End Sub
Private Function SaveRefused(ByVal id As Long, ByVal key As String, ByVal value As Variant) As Boolean
    On Error GoTo Refused
    Edit id, key, value
    Exit Function
Refused:
    SaveRefused = InStr(Err.Description, "Pull it back to edit") > 0
End Function
Private Function Refused(ByVal action As String, ByVal run As Long) As Boolean
    On Error GoTo Failed
    Select Case action
        Case "forward": ForwardRoute run, ""
        Case "pullback": PullBackRoute run, "Synthetic pull back"
        Case "sendback": SendBackRoute run, 0, "Synthetic send back"
    End Select
    Exit Function
Failed:
    Refused = True
End Function
Private Function AssignmentRefused(ByVal id As Long, ByVal reviewer As Long) As Boolean
    On Error GoTo Refused
    Call AssignReview(id, CurrentRevision(id), reviewer, Null, "Duplicate")
    Exit Function
Refused:
    AssignmentRefused = InStr(Err.Description, "already in this report's review route") > 0
End Function
Private Function StepValue(ByVal run As Long, ByVal stepNumber As Long, ByVal field As String) As Variant
    StepValue = DLookup(field, "RouteRunStep", "RunID=" & run & " AND StepOrdinal=" & stepNumber)
End Function
Private Function Events(ByVal run As Long, ByVal kind As String) As Long
    Events = DCount("*", "RouteRunEvent", "RunID=" & run & " AND EventKind='" & kind & "'")
End Function
Public Function TestRouteChain() As String
    Dim requester As Long, a As Long, b As Long, c As Long, template As Long, id As Long, run As Long, notice As String
    Dim g As Long, member As Long, other As Long, extra As Long, groupRun As Long, task As Long, assigned As Long, where As String
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC ROUTE CHAIN", "Requester"
    requester = MajorAuthor()
    a = AddContributor("Division Officer"): b = AddContributor("Department Head"): c = AddContributor("Executive Officer")
    template = CreateRouteTemplate("Chain of command", a & "," & b & "," & c, "Review the draft")
    id = CommitDraft(0, "EVAL", "Chain report", "SYNTHETIC CHAIN", #1/1/2026#, #9/30/2026#, "")
    run = StartRoute(id, template)
    Check ReviewHolder(id, where) = a And InStr(where, "step 1 of 3") > 0, "Step 1 reviewer does not hold the report: " & where
    Check StepValue(run, 1, "IncomingRevisionID") = CurrentRevision(id) And Events(run, "Start") = 1, "Start not recorded"

    ' Only the holder saves, and the holder's saves keep the route current.
    Check SaveRefused(id, "Title", "Requester edit"), "A non-holder saved a routed report"
    Check Refused("forward", run), "A non-holder forwarded"
    Check SwitchContributor(a), "Switch to step 1"
    Edit id, "Title", "Division officer title"
    Check RouteIsCurrent(run) And UXReportState(id) = "In review", "Holder save made the route stale"
    ForwardRoute run, ""
    Check DLookup("CurrentStep", "RouteRun", "RunID=" & run) = 2 And StepValue(run, 1, "OutgoingRevisionID") = CurrentRevision(id) And StepValue(run, 2, "IncomingRevisionID") = CurrentRevision(id), "Forward did not record the handoff revisions"
    Check SaveRefused(id, "Title", "Late edit"), "A forwarded reviewer saved again"
    Check SwitchContributor(b), "Switch to step 2"
    Edit id, "OfficialComments", "Department head narrative"
    ForwardRoute run, "Looks right"
    Check SwitchContributor(c), "Switch to step 3"
    Edit id, "Title", "Executive officer title"

    ' Send back: the chain runs upward again from step 1; earlier decisions stay as events.
    SendBackRoute run, 1, "Rework the title"
    Check DLookup("CurrentStep", "RouteRun", "RunID=" & run) = 1 And StepValue(run, 2, "StepState") = "Pending" And IsNull(StepValue(run, 1, "OutgoingRevisionID")), "Send back did not reset the chain"
    Check Events(run, "Forward") = 2 And Events(run, "Send back") = 1, "Send back lost the earlier decisions"
    Check Refused("sendback", run), "A non-holder sent the report back"

    ' A reviewer later in the chain may save while a lower step holds; it is flagged.
    Edit id, "OfficialComments", "Executive officer narrative"
    Check Events(run, "Higher change") = 1 And InStr(RouteFlags(run), "Changed by Executive Officer (step 3)") > 0, "Higher reviewer's change not flagged"
    Check SwitchContributor(b), "Switch to step 2 while step 1 holds"
    Edit id, "OfficialComments", "Department head narrative again"
    Check Events(run, "Higher change") = 2, "Step 2 could not save above the holder"
    ' A lower step saving over a later reviewer's change is allowed and flagged.
    Check SwitchContributor(a), "Switch back to step 1"
    Edit id, "Title", "Division officer rework"
    notice = RouteSaveNotice()
    Check InStr(notice, "Overrides a change by Executive Officer (step 3)") > 0 And Events(run, "Override") = 1, "Override not flagged: " & notice
    ForwardRoute run, ""
    Check InStr(RouteFlags(run), "Overrides a change by Executive Officer") = 0, "Flag shown after the report moved on"

    ' Pull back: an earlier reviewer, then the requester.
    Check RoutePullBackStep(run, c) = -1 And RoutePullBackStep(run, a) = 1, "Pull-back eligibility wrong"
    PullBackRoute run, "One more change"
    Check DLookup("CurrentStep", "RouteRun", "RunID=" & run) = 1 And ReviewHolder(id) = a And Events(run, "Pull back") = 1, "Pull back by step 1 failed"
    Check SwitchContributor(requester), "Switch to requester"
    PullBackRoute run, "Requester correction"
    Check DLookup("CurrentStep", "RouteRun", "RunID=" & run) = 0 And ReviewHolder(id) = requester, "Pull back by requester failed"
    Edit id, "OfficialComments", "Requester narrative"
    ForwardRoute run, "Resubmitted"
    Check DLookup("CurrentStep", "RouteRun", "RunID=" & run) = 1, "Requester resubmit did not reach step 1"

    ' Forward through every step, then local approval locks the current revision.
    Check SwitchContributor(a), "Final pass step 1": ForwardRoute run, ""
    Check SwitchContributor(b), "Final pass step 2": ForwardRoute run, ""
    Check SwitchContributor(c), "Final pass step 3": ForwardRoute run, "Approved for release"
    Check DLookup("RunState", "RouteRun", "RunID=" & run) = "Accepted" And ReviewHolder(id) = c, "Final forward did not complete the route"
    ApproveSavedRoute run, "Local approval"
    Check DLookup("ApprovedRevisionID", "RouteRun", "RunID=" & run) = CurrentRevision(id) And UXReportState(id) = "Local approval", "Approval did not lock the current revision"
    Check ReviewHolder(id) = 0, "Approved report still held"

    ' Group routes: another member's edit keeps the route current; membership changes do not.
    Check SwitchContributor(requester), "Switch to requester for groups"
    member = CommitDraft(0, "EVAL", "Group member", "SYNTHETIC MEMBER", #1/1/2026#, #9/30/2026#, "")
    other = CommitDraft(0, "EVAL", "Other member", "SYNTHETIC OTHER", #1/1/2026#, #9/30/2026#, "")
    g = CreateSummaryGroup("Chain group", "EVAL", "E5")
    AddGroupMember g, member, "E5": AddGroupMember g, other, "E5"
    groupRun = StartRoute(member, template, g)
    Edit other, "OfficialComments", "Unrelated narrative"
    Check RouteIsCurrent(groupRun), "Another member's edit staled the group route"
    ' The Members grid is outside the chain: allowed for anyone, logged on the report.
    RouteGridContext True
    Edit member, "Trait33", "4.0"
    RouteGridContext False
    Check Events(groupRun, "Grid change") = 1 And InStr(RouteFlags(groupRun), "Members grid") > 0, "Grid change not logged for the holder"
    extra = CommitDraft(0, "EVAL", "Extra member", "SYNTHETIC EXTRA", #1/1/2026#, #9/30/2026#, "")
    AddGroupMember g, extra, "E5"
    Check Not RouteIsCurrent(groupRun) And UXReportState(member) = "Review needs restart", "Membership change did not stale the group route"

    ' Assignments: the assignee holds the report and may save before deciding.
    assigned = CommitDraft(0, "EVAL", "Assigned report", "SYNTHETIC ASSIGNED", #1/1/2026#, #9/30/2026#, "")
    task = AssignReview(assigned, CurrentRevision(assigned), a, Null, "Check it")
    Check SaveRefused(assigned, "Title", "Requester edit"), "Requester saved during an assignment"
    Check SwitchContributor(a), "Switch to assignee"
    Edit assigned, "Title", "Assignee edit"
    Check FinishReview(task, "Accepted", "Edited and accepted") And DLookup("FinishedRevisionID", "ReviewTask", "TaskID=" & task) = CurrentRevision(assigned), "Assignee could not accept after editing"
    ' Someone already in the active route cannot also get an assignment; others can.
    Check SwitchContributor(requester), "Switch to requester for assignments"
    Check AssignmentRefused(member, a), "Assigned a reviewer already in the route"
    task = AssignReview(member, CurrentRevision(member), requester, Null, "Outside the route")
    Check task > 0, "Assignment to someone outside the route refused"
    Call FinishReview(task, "Cancelled", "Test cleanup")
    ' Replace swaps the active route; the old one stays as cancelled history.
    Check SwitchContributor(requester), "Switch to requester for replace"
    Dim replaced As Long, other2 As Long, f As Access.Form
    other2 = CreateRouteTemplate("Short chain", CStr(b), "Quick look")
    run = StartRoute(assigned, template)
    replaced = ReplaceRoute(run, other2)
    Check DLookup("RunState", "RouteRun", "RunID=" & run) = "Cancelled" And Events(run, "Replaced") = 1 And ReviewHolder(assigned) = b, "Replace did not swap the route"
    ' Deleting a template hides it from new reviews; the route keeps its reviewers.
    DeleteRouteTemplate other2
    Check DLookup("IsArchived", "RouteTemplate", "TemplateID=" & other2) = True And DCount("*", "RouteRunStep", "RunID=" & replaced) = 1 And ReviewHolder(assigned) = b, "Deleting a template changed a running route"
    ' Forward through the review page on a multi-step route (the path that hit error 3218).
    CancelRoute replaced
    run = StartRoute(assigned, template)
    Check SwitchContributor(a), "Switch to step 1 for the review page"
    Check UXOpenReportTool(assigned, "Review"), "Review page did not open"
    Set f = UXReviewForm(): f!txtNote = "Forward from the page"
    Check UXReviewAction("Accept") And DLookup("CurrentStep", "RouteRun", "RunID=" & run) = 2, "Review-page Forward failed: " & f!lblStatus.Caption
    Check UXLeaveEditor("Discard"), "Report close after Forward"
    TestRouteChain = "PASS: " & checks & " shared-draft review chain assertions": Exit Function
Failed:
    RouteGridContext False
    TestRouteChain = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
