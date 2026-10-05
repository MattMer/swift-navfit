Attribute VB_Name = "MajorHarness"
Option Compare Binary
Option Explicit
Private count As Long
Private Sub Check(ByVal value As Boolean, ByVal message As String)
    If Not value Then Err.Raise vbObjectError + 4900, , message
    count = count + 1
End Sub
Public Function TestMajor() As String
    Dim a As Long, b As Long, g As Long, snap As Long, meeting As Long, proposal As Long, batch As Long, route As Long, template As Long, rev As Long, failed As Boolean, undo As Long, suggestion As Object, secondProfile As Long, firstProfile As Long
    On Error GoTo Failed
    count = 0
    Call ConfigureWorkspace("SYNTHETIC COMMAND", "Synthetic contributor")
    a = CommitDraft(0, "EVAL", "Synthetic A", "Example A", Null, Null, "", "Test")
    b = CommitDraft(0, "EVAL", "Synthetic B", "Example B", Null, Null, "", "Test")
    batch = PreviewFieldBatch(a & "," & b, "OfficialComments", "O'Brien & synthetic", "Reviewed batch")
    Check Nz(MajorScalar("SELECT OfficialComments FROM Report WHERE ReportID=" & a), "") = "", "Preview does not commit values"
    ApplyFieldBatch batch
    Check MajorScalar("SELECT OfficialComments FROM Report WHERE ReportID=" & a) = "O'Brien & synthetic", "Batch writes first report"
    Check MajorScalar("SELECT OfficialComments FROM Report WHERE ReportID=" & b) = "O'Brien & synthetic", "Batch writes second report"
    Check MajorScalar("SELECT BatchState FROM ChangeBatch WHERE BatchID=" & batch) = "Applied", "Batch records completion"
    batch = PreviewFieldBatch(a & "," & b, "Title", "Reviewed title", "Stale check")
    Call CommitDraft(b, "EVAL", "Concurrent edit", "Example B", Null, Null, "", "Concurrent")
    rev = CurrentRevision(a)
    On Error Resume Next
    ApplyFieldBatch batch
    failed = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    Check failed, "Stale second report rejects entire batch"
    Check CurrentRevision(a) = rev, "Rollback restores first report revision"
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & a) = "Synthetic A", "Rollback restores first report value"
    Check MajorScalar("SELECT BatchState FROM ChangeBatch WHERE BatchID=" & batch) = "Preview", "Failed batch remains pending"
    g = CreateSummaryGroup("Synthetic E5 group", "EVAL", "E5")
    AddGroupMember g, a, "E5"
    AddGroupMember g, b, "E5"
    snap = FreezeSummaryGroup(g, "Synthetic roster checked against test list")
    Check DCount("*", "GroupSnapshotMember", "SnapshotID=" & snap) = 2, "Frozen roster contains both saved reports"
    Check MajorScalar("SELECT Stamp FROM GroupSnapshot WHERE SnapshotID=" & snap) = GroupStamp(g), "Frozen stamp matches current group"
    meeting = CreateComparisonSet(snap, "Synthetic meeting")
    proposal = ProposeMeetingField(meeting, a, "Title", "Meeting decision")
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & a) = "Synthetic A", "Proposal does not edit report"
    ApplyMeetingProposal proposal
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & a) = "Meeting decision", "Explicit proposal apply creates edit"
    Check MajorScalar("SELECT Stamp FROM GroupSnapshot WHERE SnapshotID=" & snap) <> GroupStamp(g), "Edit makes group snapshot stale"
    Check MajorScalar("SELECT RevisionID FROM ComparisonMember WHERE SetID=" & meeting & " AND ReportID=" & a) = rev, "Meeting remains bound to original revision"
    template = CreateRouteTemplate("Synthetic one-step route", CStr(MajorAuthor()), "Check saved report")
    route = StartRoute(a, template, g)
    Check RouteIsCurrent(route), "New route is current"
    DecideRouteStep route, "Accepted", "Synthetic review complete"
    Check MajorScalar("SELECT RunState FROM RouteRun WHERE RunID=" & route) = "Accepted", "Last accepted step completes route"
    Check RouteIsCurrent(route), "Accepted route remains current before edits"
    batch = PreviewFieldBatch(CStr(a), "Title", "After approval", "Material edit")
    ApplyFieldBatch batch
    Check Not RouteIsCurrent(route), "Material edit makes approval stale"
    Check MajorScalar("SELECT RunState FROM RouteRun WHERE RunID=" & route) = "Accepted", "Historical approval preserved"
    Check DCount("*", "GroupSnapshotMember", "SnapshotID=" & snap) = 2, "Frozen roster unchanged after edits"
    batch = PreviewFieldBatch(a & "," & b, "Title", "After approval", "First no-op, then change")
    rev = CurrentRevision(a): ApplyFieldBatch batch
    Check CurrentRevision(a) = rev, "No-op item creates no history"
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & b) = "After approval", "No-op item does not roll back later batch work"
    undo = PreviewBatchUndo(batch)
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & b) = "After approval", "Undo preview does not replace values"
    ApplyFieldBatch undo
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & b) = "Concurrent edit", "Undo restores original value as new revision"
    SetFieldDefault "Command", 1, "ShipStation", "COMMAND DEFAULT"
    Set suggestion = ResolveFieldDefault(a, "ShipStation")
    Check suggestion("Value") = "COMMAND DEFAULT", "Command default resolves"
    SetFieldDefault "Group", g, "ShipStation", "GROUP DEFAULT"
    Set suggestion = ResolveFieldDefault(a, "ShipStation", g)
    Check suggestion("Value") = "GROUP DEFAULT", "Group overrides command default"
    SetFieldDefault "Report", a, "ShipStation", Null
    Set suggestion = ResolveFieldDefault(a, "ShipStation", g)
    Check IsNull(suggestion("Value")), "Explicit null report default overrides inherited text"
    Check IsNull(MajorScalar("SELECT ShipStation FROM Report WHERE ReportID=" & a)), "Defaults do not silently commit report fields"
    snap = FreezeSummaryGroup(g, "Synthetic refreshed roster")
    meeting = CreateComparisonSet(snap, "Synthetic proposal batch")
    proposal = ProposeMeetingField(meeting, a, "Title", "Meeting batch A")
    undo = ProposeMeetingField(meeting, b, "Title", "Meeting batch B")
    ApplyMeetingProposals proposal & "," & undo
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & a) = "Meeting batch A", "Meeting batch changes first report"
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & b) = "Meeting batch B", "Meeting batch changes second report atomically"
    SetMeetingOrder meeting, b & "," & a
    Check MajorScalar("SELECT SortOrdinal FROM ComparisonMember WHERE SetID=" & meeting & " AND ReportID=" & b) = 1, "Manual meeting order retained"
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & b) = "Meeting batch B", "Display ordering does not assign recommendations or alter draft values"
    firstProfile = MajorAuthor(): secondProfile = AddContributor("Synthetic second reviewer")
    template = CreateRouteTemplate("Two reviewers", firstProfile & "," & secondProfile, "Independent steps")
    route = StartRoute(b, template)
    DecideRouteStep route, "Accepted", "First step complete"
    Check MajorScalar("SELECT CurrentStep FROM RouteRun WHERE RunID=" & route) = 2, "Acceptance advances to second reviewer"
    Check MajorScalar("SELECT RunState FROM RouteRun WHERE RunID=" & route) = "Pending", "Intermediate acceptance does not approve route"
    On Error Resume Next
    DecideRouteStep route, "Accepted", "Wrong reviewer"
    failed = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Check failed, "Wrong contributor cannot decide second step"
    Call SwitchContributor(secondProfile)
    DecideRouteStep route, "Returned", "Needs correction"
    Check MajorScalar("SELECT RunState FROM RouteRun WHERE RunID=" & route) = "Returned", "Assigned second reviewer can return route"
    Call SwitchContributor(firstProfile)
    template = CreateRouteTemplate("Repeated reviewer", firstProfile & "," & firstProfile, "Two distinct steps")
    Check DCount("*", "RouteTemplateStep", "TemplateID=" & template) = 2, "Same reviewer can appear at different steps"
    route = StartRoute(b, template, g)
    DecideRouteStep route, "Accepted", "First repeated step"
    DecideRouteStep route, "Accepted", "Final repeated step"
    ApproveSavedRoute route, "Synthetic local approval"
    Check ApprovalIsCurrent(route), "Local approval is current after locking"
    Check MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & b) = "Approved", "Local approval locks report editing"
    rev = CurrentRevision(b)
    On Error Resume Next
    Call CommitDraft(b, "EVAL", "Forbidden locked edit", "Example B", Null, Null, "", "Locked test")
    failed = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Check failed And CurrentRevision(b) = rev, "Locked edit rejected without history"
    RemoveGroupMember g, a
    Check Not ApprovalIsCurrent(route), "Membership change makes local approval stale"
    ReopenApprovedReport b, "Synthetic correction"
    Check MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & b) = "Draft", "Explicit reopen restores editing"
    Check MajorScalar("SELECT RunState FROM RouteRun WHERE RunID=" & route) = "Reopened", "Reopening retains approval history"
    Check DCount("*", "ReportEvent", "ReportID=" & b) = 2, "Approval and reopening have contributor events"
    TestMajor = "PASS: " & count & " major workflow assertions"
    Exit Function
Failed:
    TestMajor = "FAIL after " & count & " assertions: " & Err.Number & ": " & Err.Description
End Function
