Attribute VB_Name = "modUXReview"
Option Compare Binary
Option Explicit
Public Function UXRouteLabel(ByVal runID As Long) As String
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & runID, dbOpenSnapshot)
    If Not r.EOF Then UXRouteLabel = UXMemberLabel(r!ReportID) & " | " & Nz(DLookup("Title", "RouteTemplate", "TemplateID=" & r!TemplateID), "Review route") & " | Saved revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & r!RevisionID), "?") & " | " & r!RunState & " | " & Format$(r!RequestedAt, "yyyy-mm-dd hh:nn")
    r.Close
End Function
Public Function UXLoadReview() As Boolean
    Dim f As Access.Form, id As Long, r As DAO.Recordset, args As Variant
    Set f = Forms!frmUXReview: id = Val(Nz(f.OpenArgs, "0")): f!txtReportID = id
    If InStr(Nz(f.OpenArgs, ""), ":") > 0 Then f!txtGroupID = Val(Split(f.OpenArgs, ":")(1)) Else f!txtGroupID = 0
    Set r = CurrentDb.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    If r.EOF Then r.Close: Exit Function
    f!lblHeading.Caption = "Review - " & UXMemberLabel(id)
    f!lblSaved.Caption = r!FormKind & " | " & Nz(r!PeriodStart, "No start date") & " to " & Nz(r!PeriodEnd, "No end date") & " | Current saved revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & r!CurrentRevisionID), "?")
    r.Close
    f!cboTemplate.RowSource = "SELECT TemplateID,Title FROM RouteTemplate WHERE IsArchived=False ORDER BY Title"
    f!cboRun.RowSource = "SELECT RunID,UXRouteLabel(RunID) AS Description FROM RouteRun WHERE ReportID=" & id & " ORDER BY RunID DESC"
    f!cboRun.Requery
    args = Split(Nz(f.OpenArgs, "0"), ":")
    If UBound(args) >= 2 And IsNull(f!cboRun) Then
        If DCount("*", "RouteRun", "RunID=" & Val(args(2)) & " AND ReportID=" & id) > 0 Then f!cboRun = Val(args(2))
    End If
    If IsNull(f!cboRun) And f!cboRun.ListCount > 0 Then f!cboRun = f!cboRun.ItemData(0)
    Call UXReviewLayout
    UXLoadReview = UXReviewStatus()
End Function
Public Function UXReviewStatus() As Boolean
    Dim f As Access.Form, id As Long, run As Long, r As DAO.Recordset, state As String, current As Boolean, author As Long, assignee As Long, editable As Boolean, message As String
    Set f = Forms!frmUXReview: id = Nz(f!txtReportID, 0): run = Nz(f!cboRun, 0): author = MajorAuthor()
    editable = Not Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True) And Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") = "Draft"
    f!cmdStart.Enabled = editable And DCount("*", "RouteRun", "ReportID=" & id & " AND RunState='Pending'") = 0
    If Nz(f!txtGroupID, 0) > 0 Then f!cmdStart.Enabled = f!cmdStart.Enabled And Not Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & f!txtGroupID), True)
    f!cmdAccept.Enabled = False: f!cmdReturn.Enabled = False: f!cmdRestart.Enabled = False: f!cmdApprove.Enabled = False: f!cmdCancel.Enabled = False
    f!cmdReopen.Enabled = Not Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True) And Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") = "Approved"
    f!lstSteps.RowSource = "": f!txtStep = ""
    message = "Choose a route template and Start review. Actions use saved values; pending editor changes remain separate."
    If Not editable Then message = "Editing is locked. Reopen local approval with a reason, or recover the archived report first."
    If run > 0 Then
        Set r = CurrentDb.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & run & " AND ReportID=" & id, dbOpenSnapshot)
        If r.EOF Then r.Close: Err.Raise vbObjectError + 5901, , "Choose a route for this report."
        state = r!RunState: current = RouteIsCurrent(run)
        f!txtExpectedRevision = CurrentRevision(id): f!txtExpectedStamp = ""
        If Nz(r!GroupID, 0) > 0 Then f!txtExpectedStamp = GroupStamp(r!GroupID)
        assignee = Nz(DLookup("AssigneeID", "RouteRunStep", "RunID=" & run & " AND StepOrdinal=" & r!CurrentStep), 0)
        f!lstSteps.RowSource = "SELECT StepOrdinal AS [Step],DLookup('DisplayName','UserProfile','UserProfileID=' & AssigneeID) AS Reviewer,StepState AS [Status],Instructions,DecisionNote AS [Decision note] FROM RouteRunStep WHERE RunID=" & run & " ORDER BY StepOrdinal"
        f!lstSteps.Requery
        message = "Route: " & state & ". Saved revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & r!RevisionID), "?") & ". "
        If Not current And (state = "Pending" Or state = "Accepted") Then
            message = message & "Needs restart: saved report values or group membership changed. The requester can restart using the same template; all earlier decisions stay in history."
            f!cmdRestart.Enabled = editable And r!RequestedBy = author
            If Nz(r!GroupID, 0) > 0 Then f!cmdRestart.Enabled = f!cmdRestart.Enabled And Not Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & r!GroupID), True)
        ElseIf state = "Pending" Then
            message = message & "Current step " & r!CurrentStep & " of " & DCount("*", "RouteRunStep", "RunID=" & run) & ": " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & assignee), "Unknown reviewer") & ". "
            If assignee = author Then message = message & "Enter a decision note, then Accept step or Return report." Else message = message & "Switch to this assigned contributor to record a decision."
            f!cmdAccept.Enabled = current And editable And assignee = author: f!cmdReturn.Enabled = f!cmdAccept.Enabled
        ElseIf state = "Accepted" And current Then
            message = message & "All steps accepted. The final reviewer can record Local approval with a note."
            f!cmdApprove.Enabled = editable And Nz(r!ApprovedBy, 0) = author
        Else
            message = message & "This route is history. Start another review if needed; local approval must be reopened before editing."
        End If
        f!cmdCancel.Enabled = state = "Pending" And r!RequestedBy = author
        r.Close
    End If
    f!cboRun.Requery
    f!cboTemplate.Requery
    f!lblStatus.Caption = message
    f!lblUnavailable.Caption = "Disabled actions: decisions require the assigned contributor and a current pending route; restart/cancel require its requester; approval requires the final reviewer. Archived reports and groups must be recovered. Start requires a draft without a pending route."
    UXReviewStatus = True
End Function
Public Function UXReviewStep() As Boolean
    Dim f As Access.Form, r As DAO.Recordset
    Set f = Forms!frmUXReview
    If IsNull(f!lstSteps) Then Exit Function
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RouteRunStep WHERE RunID=" & Nz(f!cboRun, 0) & " AND StepOrdinal=" & CLng(f!lstSteps), dbOpenSnapshot)
    If Not r.EOF Then f!txtStep = "Reviewer: " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & r!AssigneeID), "Unknown") & vbCrLf & "Instructions: " & Nz(r!Instructions, "None") & vbCrLf & "Decision: " & r!StepState & vbCrLf & Nz(r!DecisionNote, "No decision note") & vbCrLf & Nz(r!DecisionAt, "")
    r.Close: UXReviewStep = True
End Function
Public Function RestartStaleRoute(ByVal runID As Long, ByVal expectedRevision As Long, ByVal expectedStamp As String) As Long
    Dim w As DAO.Workspace, db As DAO.Database, r As DAO.Recordset, id As Long, template As Long, group As Long, failure As Long, message As String, trans As Boolean, replacement As Long
    On Error GoTo Failed
    Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & runID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 5902, , "Route no longer exists."
    id = r!ReportID: template = r!TemplateID: group = Nz(r!GroupID, 0)
    If r!RequestedBy <> MajorAuthor() Then Err.Raise vbObjectError + 5902, , "Only the requester can restart this route."
    If (r!RunState <> "Pending" And r!RunState <> "Accepted") Or RouteIsCurrent(runID) Then Err.Raise vbObjectError + 5902, , "Only a stale pending or accepted route can be restarted."
    If CurrentRevision(id) <> expectedRevision Then Err.Raise vbObjectError + 5902, , "Saved report changed; refresh review before restarting."
    If group > 0 Then
        If GroupStamp(group) <> expectedStamp Or Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & group), True) Then Err.Raise vbObjectError + 5902, , "Group changed or is archived; refresh or recover it before restarting."
    End If
    If Nz(DLookup("IsArchived", "RouteTemplate", "TemplateID=" & template), True) Then Err.Raise vbObjectError + 5902, , "Recover the route template before restarting."
    r.Close: Set r = Nothing
    Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    If CurrentRevision(id) <> expectedRevision Then Err.Raise vbObjectError + 5902, , "Saved report changed; refresh review before restarting."
    If group > 0 Then
        If GroupStamp(group) <> expectedStamp Then Err.Raise vbObjectError + 5902, , "Group changed; refresh review before restarting."
    End If
    db.Execute "UPDATE RouteRun SET RunState='Cancelled' WHERE RunID=" & runID, dbFailOnError
    replacement = StartRoute(id, template, group, False)
    w.CommitTrans: trans = False
    RestartStaleRoute = replacement: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not r Is Nothing Then r.Close
    If trans Then w.Rollback
    On Error GoTo 0
    Err.Raise failure, "RestartStaleRoute", message
End Function
Public Function UXReviewAction(ByVal action As String) As Boolean
    Dim f As Access.Form, run As Long, id As Long, group As Long, stamp As String, expected As Long
    On Error GoTo Failed
    Set f = Forms!frmUXReview: run = Nz(f!cboRun, 0): id = Nz(f!txtReportID, 0)
    Select Case action
        Case "Start": f!cboRun = StartRoute(id, Nz(f!cboTemplate, 0), Nz(f!txtGroupID, 0))
        Case "Restart"
            expected = Nz(f!txtExpectedRevision, 0): stamp = Nz(f!txtExpectedStamp, "")
            id = RestartStaleRoute(run, expected, stamp)
            f!cboRun.Requery: f!cboRun = id
        Case "Accept": DecideRouteStep run, "Accepted", Nz(f!txtNote, "")
        Case "Return": DecideRouteStep run, "Returned", Nz(f!txtNote, "")
        Case "Cancel": CancelRoute run
        Case "Approve": ApproveSavedRoute run, Nz(f!txtNote, "")
        Case "Reopen": ReopenApprovedReport id, Nz(f!txtNote, "")
        Case Else: Err.Raise vbObjectError + 5903, , "Unknown review action."
    End Select
    Call UXRefreshEditorWorkflow(CLng(f!txtReportID))
    f!cmdBack.SetFocus
    f!txtNote = Null
    Call UXLoadReview
    Call UXWorkChanged
    UXReviewAction = True: Exit Function
Failed:
    If CurrentProject.AllForms("frmUXReview").IsLoaded Then Forms!frmUXReview!lblStatus.Caption = Err.Description
End Function
Public Function UXGroupReview() As Boolean
    Dim i As Long, id As Long, f As Access.Form
    Set f = Forms!frmMajorWorkbench
    For i = 0 To f!lstReports.ListCount - 1
        If f!lstReports.Selected(i) Then
            If id > 0 Then f!lblStatus.Caption = "Select one member to review.": Exit Function
            id = CLng(f!lstReports.Column(0, i))
        End If
    Next
    If id = 0 Then f!lblStatus.Caption = "Select one member to review.": Exit Function
    If CurrentProject.AllForms("frmUXReview").IsLoaded Then DoCmd.Close acForm, "frmUXReview", acSaveNo
    DoCmd.OpenForm "frmUXReview", , , , , , CStr(id) & ":" & Nz(f!cboGroup, 0)
    UXGroupReview = True
End Function















