Attribute VB_Name = "modRouteTemplates"
Option Compare Binary
Option Explicit
Public Function CreateRouteTemplate(ByVal title As String, ByVal assigneesCSV As String, ByVal instructions As String) As Long
    Dim ids As Object, key As Variant, r As DAO.Recordset, s As DAO.Recordset, id As Long, ordinal As Long, w As DAO.Workspace, failure As Long, message As String, operation As Long, author As Long
    author = MajorAuthor()
    If Len(Trim$(title)) = 0 Or Len(title) > 120 Then Err.Raise vbObjectError + 4300, , "Enter a route title."
    If Len(Trim$(assigneesCSV)) = 0 Then Err.Raise vbObjectError + 4300, , "Choose at least one reviewer."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("RouteTemplate", dbOpenDynaset)
    r.AddNew: r!Title = title: r!IsArchived = False: r.Update: r.Bookmark = r.LastModified: id = r!TemplateID: r.Close
    Set s = CurrentDb.OpenRecordset("RouteTemplateStep", dbOpenDynaset)
    For Each key In Split(assigneesCSV, ",")
        Set ids = MajorIDs(CStr(key))
        ordinal = ordinal + 1
        s.AddNew: s!TemplateID = id: s!StepOrdinal = ordinal: s!AssigneeID = CLng(Trim$(key)): s!Instructions = instructions: s.Update
    Next
    s.Close
    operation = BeginOperation("Create route template", "Template " & id & ": " & title & "; ordered contributors: " & assigneesCSV & "; instructions: " & instructions)
    FinishOperation operation, "Completed", "Template stored; future routes copy its steps."
    w.CommitTrans: CreateRouteTemplate = id: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not r Is Nothing Then r.Close
    If Not s Is Nothing Then s.Close
    On Error GoTo 0
    Err.Raise failure, "CreateRouteTemplate", message
End Function
Public Function StartRoute(ByVal reportID As Long, ByVal templateID As Long, Optional ByVal groupID As Long = 0) As Long
    Dim w As DAO.Workspace, r As DAO.Recordset, s As DAO.Recordset, t As DAO.Recordset, id As Long, failure As Long, message As String
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & reportID), True) Then Err.Raise vbObjectError + 4301, , "Select an active report."
    If Nz(MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & reportID), "") <> "Draft" Then Err.Raise vbObjectError + 4301, , "Reopen the approved report before starting another route."
    If DCount("*", "RouteRun", "ReportID=" & reportID & " AND RunState='Pending'") > 0 Then Err.Raise vbObjectError + 4302, , "Finish or cancel the active route first."
    If Nz(MajorScalar("SELECT IsArchived FROM RouteTemplate WHERE TemplateID=" & templateID), True) Then Err.Raise vbObjectError + 4303, , "Select an active template."
    Set t = CurrentDb.OpenRecordset("SELECT * FROM RouteTemplateStep WHERE TemplateID=" & templateID & " ORDER BY StepOrdinal", dbOpenSnapshot)
    If t.EOF Then Err.Raise vbObjectError + 4304, , "Route template is empty."
    Set r = CurrentDb.OpenRecordset("RouteRun", dbOpenDynaset)
    r.AddNew: r!ReportID = reportID: r!RevisionID = CurrentRevision(reportID): r!TemplateID = templateID: r!CurrentStep = 1: r!RunState = "Pending": r!RequestedBy = MajorAuthor(): r!RequestedAt = Now()
    If groupID > 0 Then
        If DCount("*", "GroupMember", "GroupID=" & groupID & " AND ReportID=" & reportID) <> 1 Then Err.Raise vbObjectError + 4305, , "Report is not in the selected group."
        r!GroupID = groupID: r!GroupStamp = GroupStamp(groupID)
    End If
    r.Update: r.Bookmark = r.LastModified: id = r!RunID: r.Close
    Set s = CurrentDb.OpenRecordset("RouteRunStep", dbOpenDynaset)
    Do While Not t.EOF
        s.AddNew: s!RunID = id: s!StepOrdinal = t!StepOrdinal: s!AssigneeID = t!AssigneeID: s!Instructions = t!Instructions: s!StepState = "Pending": s.Update: t.MoveNext
    Loop
    s.Close: t.Close: w.CommitTrans: StartRoute = id: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not r Is Nothing Then r.Close
    If Not s Is Nothing Then s.Close
    If Not t Is Nothing Then t.Close
    On Error GoTo 0
    Err.Raise failure, "StartRoute", message
End Function
Public Function RouteIsCurrent(ByVal runID As Long) As Boolean
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & runID, dbOpenSnapshot)
    If r.EOF Then r.Close: Exit Function
    RouteIsCurrent = (r!RevisionID = CurrentRevision(r!ReportID))
    If Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & r!ReportID), True) Then RouteIsCurrent = False
    If Not IsNull(r!GroupID) Then
        If StrComp(Nz(r!GroupStamp, ""), GroupStamp(r!GroupID), vbBinaryCompare) <> 0 Then RouteIsCurrent = False
    End If
    r.Close
End Function
Public Sub DecideRouteStep(ByVal runID As Long, ByVal decision As String, ByVal note As String)
    Dim r As DAO.Recordset, s As DAO.Recordset, w As DAO.Workspace, failure As Long, message As String
    If decision <> "Accepted" And decision <> "Returned" Then Err.Raise vbObjectError + 4306, , "Choose Accepted or Returned."
    If Len(Trim$(note)) = 0 Then Err.Raise vbObjectError + 4307, , "Record a decision note."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & runID, dbOpenDynaset)
    If r.EOF Then Err.Raise vbObjectError + 4308, , "Route does not exist."
    If r!RunState <> "Pending" Or Not RouteIsCurrent(runID) Then Err.Raise vbObjectError + 4309, , "Route is finished or stale; start a new route for current values."
    Set s = CurrentDb.OpenRecordset("SELECT * FROM RouteRunStep WHERE RunID=" & runID & " AND StepOrdinal=" & r!CurrentStep, dbOpenDynaset)
    If s.EOF Then Err.Raise vbObjectError + 4310, , "Route step missing."
    If s!AssigneeID <> MajorAuthor() Then Err.Raise vbObjectError + 4311, , "Switch to the assigned contributor to record this decision."
    s.Edit: s!StepState = decision: s!DecisionNote = note: s!DecisionAt = Now(): s.Update: s.Close
    r.Edit
    If decision = "Returned" Then
        r!RunState = "Returned"
    ElseIf DCount("*", "RouteRunStep", "RunID=" & runID & " AND StepOrdinal>" & r!CurrentStep) = 0 Then
        r!RunState = "Accepted": r!ApprovedBy = MajorAuthor(): r!ApprovedAt = Now()
    Else
        r!CurrentStep = r!CurrentStep + 1
    End If
    r.Update: r.Close: w.CommitTrans: Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not r Is Nothing Then r.Close
    If Not s Is Nothing Then s.Close
    On Error GoTo 0
    Err.Raise failure, "DecideRouteStep", message
End Sub
Public Sub CancelRoute(ByVal runID As Long)
    If Nz(MajorScalar("SELECT RequestedBy FROM RouteRun WHERE RunID=" & runID), 0) <> MajorAuthor() Then Err.Raise vbObjectError + 4312, , "Only the requester can cancel this route."
    CurrentDb.Execute "UPDATE RouteRun SET RunState='Cancelled' WHERE RunID=" & runID & " AND RunState='Pending'", dbFailOnError
End Sub
Public Function ApprovalIsCurrent(ByVal runID As Long) As Boolean
    If Nz(MajorScalar("SELECT RunState FROM RouteRun WHERE RunID=" & runID), "") <> "Approved" Then Exit Function
    ApprovalIsCurrent = RouteIsCurrent(runID)
End Function
Public Sub ApproveSavedRoute(ByVal runID As Long, ByVal note As String)
    Dim r As DAO.Recordset, e As DAO.Recordset, w As DAO.Workspace, operation As Long, failure As Long, message As String
    If Len(Trim$(note)) = 0 Then Err.Raise vbObjectError + 4313, , "Record a local approval note."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & runID, dbOpenDynaset)
    If r.EOF Then Err.Raise vbObjectError + 4308, , "Route does not exist."
    If r!RunState <> "Accepted" Or Not RouteIsCurrent(runID) Then Err.Raise vbObjectError + 4314, , "Complete a current route first."
    If Nz(r!ApprovedBy, 0) <> MajorAuthor() Then Err.Raise vbObjectError + 4315, , "The final assigned reviewer must record local approval."
    If Nz(MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & r!ReportID), "") <> "Draft" Then Err.Raise vbObjectError + 4316, , "Report is already locked."
    operation = BeginOperation("Local approval", "Route " & runID & ": " & note)
    CurrentDb.Execute "UPDATE Report SET WorkflowState='Approved' WHERE ReportID=" & r!ReportID, dbFailOnError
    Set e = CurrentDb.OpenRecordset("ReportEvent", dbOpenDynaset)
    e.AddNew: e!ReportID = r!ReportID: e!AuthorProfileID = MajorAuthor(): e!EventAt = Now(): e!EventKind = "Local approval": e.Update: e.Close
    r.Edit: r!RunState = "Approved": r!ApprovedAt = Now(): r.Update: r.Close
    FinishOperation operation, "Completed", "Saved revision locally approved and editing locked; no official signature or submission."
    w.CommitTrans: Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not r Is Nothing Then r.Close
    If Not e Is Nothing Then e.Close
    On Error GoTo 0
    Err.Raise failure, "ApproveSavedRoute", message
End Sub
Public Sub ReopenApprovedReport(ByVal reportID As Long, ByVal note As String)
    Dim w As DAO.Workspace, e As DAO.Recordset, operation As Long, failure As Long, message As String
    If Len(Trim$(note)) = 0 Then Err.Raise vbObjectError + 4317, , "Record a reopening reason."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If Nz(MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & reportID), "") <> "Approved" Then Err.Raise vbObjectError + 4318, , "Select a locally approved report."
    If Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & reportID), True) Then Err.Raise vbObjectError + 4318, , "Recover the archived report first."
    operation = BeginOperation("Reopen approval", "Report " & reportID & ": " & note)
    CurrentDb.Execute "UPDATE RouteRun SET RunState='Reopened' WHERE ReportID=" & reportID & " AND RunState='Approved'", dbFailOnError
    CurrentDb.Execute "UPDATE Report SET WorkflowState='Draft' WHERE ReportID=" & reportID, dbFailOnError
    Set e = CurrentDb.OpenRecordset("ReportEvent", dbOpenDynaset)
    e.AddNew: e!ReportID = reportID: e!AuthorProfileID = MajorAuthor(): e!EventAt = Now(): e!EventKind = "Approval reopened": e.Update: e.Close
    FinishOperation operation, "Completed", "Editing reopened; previous local approval retained as history."
    w.CommitTrans: Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not e Is Nothing Then e.Close
    On Error GoTo 0
    Err.Raise failure, "ReopenApprovedReport", message
End Sub
