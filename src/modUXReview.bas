Attribute VB_Name = "modUXReview"
Option Compare Binary
Option Explicit
Public Function UXRouteLabel(ByVal runID As Long) As String
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & runID, dbOpenSnapshot)
    If Not r.EOF Then UXRouteLabel = UXMemberLabel(r!ReportID) & " | " & Nz(DLookup("Title", "RouteTemplate", "TemplateID=" & r!TemplateID), "Review route") & " | Started on revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & r!RevisionID), "?") & " | " & r!RunState & " | " & Format$(r!RequestedAt, "yyyy-mm-dd hh:nn")
    r.Close
End Function
Public Function UXLoadReview(Optional ByVal reportID As Long = 0, Optional ByVal groupID As Long = -1, Optional ByVal runID As Long = 0) As Boolean
    Dim f As Access.Form, id As Long, r As DAO.Recordset, args As Variant
    Set f = UXReviewForm()
    If reportID > 0 Then f!txtReportID = reportID
    id = Nz(f!txtReportID, Nz(Forms!frmEditor!txtReportID, 0))
    If groupID >= 0 Then f!txtGroupID = groupID
    If id = 0 Then Exit Function
    Set r = CurrentDb.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    If r.EOF Then r.Close: Exit Function
    f!lblHeading.Caption = "Review - " & UXMemberLabel(id)
    f!lblSaved.Caption = r!FormKind & " | " & Nz(r!PeriodStart, "No start date") & " to " & Nz(r!PeriodEnd, "No end date") & " | Current saved revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & r!CurrentRevisionID), "?")
    r.Close
    f!cboTemplate.RowSource = "SELECT TemplateID,Title FROM RouteTemplate WHERE IsArchived=False ORDER BY Title"
    f!cboRun.RowSource = "SELECT RunID,UXRouteLabel(RunID) AS Description FROM RouteRun WHERE ReportID=" & id & " ORDER BY RunID DESC"
    f!cboRun.Requery
    If runID > 0 Then
        If DCount("*", "RouteRun", "RunID=" & runID & " AND ReportID=" & id) > 0 Then f!cboRun = runID
    End If
    If IsNull(f!cboRun) And f!cboRun.ListCount > 0 Then f!cboRun = f!cboRun.ItemData(0)

    UXLoadReview = UXReviewStatus()
End Function
Public Function UXReviewStatus() As Boolean
    Dim f As Access.Form, id As Long, run As Long, r As DAO.Recordset, state As String, current As Boolean, author As Long, assignee As Long, editable As Boolean, message As String, primary As String, nextAction As String, action As Variant
    Set f = UXReviewForm(): id = Nz(f!txtReportID, 0): run = Nz(f!cboRun, 0): author = MajorAuthor()
    editable = Not Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True) And Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") = "Draft"
    'With an active route, Start replaces it with the selected template.
    f!cmdStart.Enabled = editable
    If DCount("*", "RouteRun", "ReportID=" & id & " AND RunState In ('Pending','Accepted')") > 0 Then f!cmdStart.Caption = "Replace current route with this template" Else f!cmdStart.Caption = "Start review of current saved revision"
    If Nz(f!txtGroupID, 0) > 0 Then f!cmdStart.Enabled = f!cmdStart.Enabled And Not Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & f!txtGroupID), True)
    f!cmdAccept.Enabled = False: f!cmdReturn.Enabled = False: f!cmdPullBack.Enabled = False: f!cmdRestart.Enabled = False: f!cmdApprove.Enabled = False: f!cmdCancel.Enabled = False
    f!cmdReopen.Enabled = Not Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True) And Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") = "Approved"
    f!lstSteps.RowSource = "": f!txtStep = ""
    message = "Choose a route template and Start review. The reviewer holding the report can edit it, then forward it or send it back."
    If Not editable Then message = "Editing is locked. Reopen local approval with a reason, or recover the archived report first."
    If run > 0 Then
        Set r = CurrentDb.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & run & " AND ReportID=" & id, dbOpenSnapshot)
        If r.EOF Then r.Close: Err.Raise vbObjectError + 5901, "modUXReview.UXReviewStatus", "Choose a route for this report."
        state = r!RunState: current = RouteIsCurrent(run)
        f!txtExpectedRevision = CurrentRevision(id): f!txtExpectedStamp = ""
        If Nz(r!GroupID, 0) > 0 Then f!txtExpectedStamp = GroupStamp(r!GroupID)
        assignee = RouteHolder(run)
        f!lstSteps.RowSource = "SELECT StepOrdinal AS [Step],DLookup('DisplayName','UserProfile','UserProfileID=' & AssigneeID) AS Reviewer,StepState AS [Status],Instructions,DecisionNote AS [Decision note] FROM RouteRunStep WHERE RunID=" & run & " ORDER BY StepOrdinal"
        f!lstSteps.Requery
        message = "Route: " & state & ". "
        If (state = "Pending" Or state = "Accepted") And current Then message = message & RouteFlags(run)
        If Not current And (state = "Pending" Or state = "Accepted") Then
            message = message & "Needs restart: the group's membership changed. The requester can restart using the same template; all earlier decisions stay in history."
            f!cmdRestart.Enabled = editable And r!RequestedBy = author
            If Nz(r!GroupID, 0) > 0 Then f!cmdRestart.Enabled = f!cmdRestart.Enabled And Not Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & r!GroupID), True)
        ElseIf state = "Pending" Then
            If r!CurrentStep = 0 Then
                message = message & "Back with the requester, " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & assignee), "Unknown contributor") & ". "
            Else
                message = message & "Step " & r!CurrentStep & " of " & DCount("*", "RouteRunStep", "RunID=" & run) & ": with " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & assignee), "Unknown reviewer") & ". "
            End If
            If assignee = author Then message = message & "You hold this report: edit it if needed, then Forward, or Send back with a note (to the step selected below, or the previous step)." Else message = message & "Only the contributor holding it can edit or forward it."
            f!cmdAccept.Enabled = current And editable And assignee = author
            f!cmdReturn.Enabled = f!cmdAccept.Enabled And r!CurrentStep > 0
        ElseIf state = "Accepted" And current Then
            message = message & "Every step has forwarded. The final reviewer can record Local approval, or Send back with a note."
            f!cmdApprove.Enabled = editable And Nz(r!ApprovedBy, 0) = author
            f!cmdReturn.Enabled = f!cmdApprove.Enabled
        Else
            message = message & "This route is history. Start another review if needed; local approval must be reopened before editing."
        End If
        f!cmdCancel.Enabled = state = "Pending" And r!RequestedBy = author
        f!cmdPullBack.Enabled = current And editable And RoutePullBackStep(run, author) >= 0
        r.Close
    End If
    If Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True) Then
        nextAction = "Recover this archived report from Reports before continuing."
    ElseIf Nz(f!txtGroupID, 0) > 0 And Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & Nz(f!txtGroupID, 0)), False) Then
        nextAction = "Recover this archived group before starting or restarting review."
    ElseIf f!cmdReopen.Enabled Then
        primary = "cmdReopen": nextAction = "Reopen local approval to edit."
    ElseIf f!cmdRestart.Enabled Then
        primary = "cmdRestart": nextAction = "Restart this stale route as its requester."
    ElseIf f!cmdAccept.Enabled Then
        primary = "cmdAccept": nextAction = "Edit if needed, then Forward, or Send back with a note."
    ElseIf f!cmdApprove.Enabled Then
        primary = "cmdApprove": nextAction = "Enter a note and record local approval as the final reviewer."
    ElseIf state = "Pending" Or state = "Accepted" Then
        nextAction = "Wait for the assigned contributor or requester shown above."
    ElseIf f!cmdStart.Enabled Then
        primary = "cmdStart": nextAction = "Choose a route template, then Start review."
    Else
        nextAction = "Recover the archived report/group or reopen local approval before continuing."
    End If
    message = "Next: " & nextAction & vbCrLf & message
    For Each action In Array("cmdStart", "cmdAccept", "cmdReturn", "cmdPullBack", "cmdRestart", "cmdApprove", "cmdCancel", "cmdReopen")
        With f.Controls(action)
            .UseTheme = False: .FontBold = (.Name = primary And .Enabled)
            .BackColor = IIf(.FontBold, UXColor("Primary"), UXColor("Neutral"))
            If .Enabled Then
                .ControlTipText = "Applies to saved values. " & nextAction
            Else
                Select Case CStr(action)
                    Case "cmdStart": .ControlTipText = "Unavailable: start or replace requires an active draft and group."
                    Case "cmdAccept", "cmdReturn": .ControlTipText = Left$("Unavailable: only the contributor holding the report can forward it or send it back. " & message, 255)
                    Case "cmdPullBack": .ControlTipText = "Unavailable: the requester or a reviewer at an earlier step can pull the report back."
                    Case "cmdApprove": .ControlTipText = "Unavailable: local approval requires all steps accepted on current values and the final reviewer."
                    Case "cmdRestart": .ControlTipText = "Unavailable: only the requester can restart a route whose group membership changed."
                    Case "cmdCancel": .ControlTipText = "Unavailable: only the requester can cancel a pending route."
                    Case "cmdReopen": .ControlTipText = "Unavailable: reopen requires an active, locally approved report."
                End Select
            End If
        End With
    Next
    f!cboRun.Requery
    f!cboTemplate.Requery
    f!lblStatus.Caption = message
    f!lblUnavailable.Caption = "Local approval is not official submission; group status is separate. Disabled actions: decisions require the assigned contributor and a current pending route; restart/cancel require its requester; approval requires the final reviewer. Archived reports and groups must be recovered. Start requires a draft without a pending route."
    UXApplyVisual f
    UXReviewStatus = True
End Function
Public Function UXReviewStep() As Boolean
    Dim f As Access.Form, r As DAO.Recordset
    Set f = UXReviewForm()
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
    If r.EOF Then Err.Raise vbObjectError + 5902, "modUXReview.RestartStaleRoute", "Route no longer exists."
    id = r!ReportID: template = r!TemplateID: group = Nz(r!GroupID, 0)
    If r!RequestedBy <> MajorAuthor() Then Err.Raise vbObjectError + 5902, "modUXReview.RestartStaleRoute", "Only the requester can restart this route."
    If (r!RunState <> "Pending" And r!RunState <> "Accepted") Or RouteIsCurrent(runID) Then Err.Raise vbObjectError + 5902, "modUXReview.RestartStaleRoute", "Only a stale pending or accepted route can be restarted."
    If CurrentRevision(id) <> expectedRevision Then Err.Raise vbObjectError + 5902, "modUXReview.RestartStaleRoute", "Saved report changed; refresh review before restarting."
    If group > 0 Then
        If GroupStamp(group) <> expectedStamp Or Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & group), True) Then Err.Raise vbObjectError + 5902, "modUXReview.RestartStaleRoute", "Group changed or is archived; refresh or recover it before restarting."
    End If
    If Nz(DLookup("IsArchived", "RouteTemplate", "TemplateID=" & template), True) Then Err.Raise vbObjectError + 5902, "modUXReview.RestartStaleRoute", "Recover the route template before restarting."
    r.Close: Set r = Nothing
    Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    If CurrentRevision(id) <> expectedRevision Then Err.Raise vbObjectError + 5902, "modUXReview.RestartStaleRoute", "Saved report changed; refresh review before restarting."
    If group > 0 Then
        If GroupStamp(group) <> expectedStamp Then Err.Raise vbObjectError + 5902, "modUXReview.RestartStaleRoute", "Group changed; refresh review before restarting."
    End If
    db.Execute "UPDATE RouteRun SET RunState='Cancelled' WHERE RunID=" & runID, dbFailOnError
    replacement = StartRoute(id, template, group, False)
    w.CommitTrans: trans = False
    RestartStaleRoute = replacement: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modUXReview", "RestartStaleRoute", "Execute", Err.Number, Err.Description, Err.Source, IIf(trans, "Active", "Not active"))
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    If trans Then SwiftRollback w, swiftPrimary
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
Public Function UXReviewAction(ByVal action As String) As Boolean
    Dim f As Access.Form, run As Long, id As Long, group As Long, stamp As String, expected As Long
    On Error GoTo Failed
    UXFlushSurface
    If EditorHasPending() And action <> "Cancel" Then Err.Raise 5, "modUXReview.UXReviewAction", "Save or undo pending report changes before recording a review action."
    Set f = UXReviewForm(): run = Nz(f!cboRun, 0): id = Nz(f!txtReportID, 0)
    Select Case action
        Case "Start"
            run = Nz(MajorScalar("SELECT Max(RunID) FROM RouteRun WHERE ReportID=" & id & " AND RunState In ('Pending','Accepted')"), 0)
            If run > 0 Then f!cboRun = ReplaceRoute(run, Nz(f!cboTemplate, 0)) Else f!cboRun = StartRoute(id, Nz(f!cboTemplate, 0), Nz(f!txtGroupID, 0))
        Case "Restart"
            expected = Nz(f!txtExpectedRevision, 0): stamp = Nz(f!txtExpectedStamp, "")
            id = RestartStaleRoute(run, expected, stamp)
            f!cboRun.Requery: f!cboRun = id
        Case "Accept": ForwardRoute run, Nz(f!txtNote, "")
        Case "Return": SendBackRoute run, UXSendBackTarget(run), Nz(f!txtNote, "")
        Case "PullBack": PullBackRoute run, Nz(f!txtNote, "")
        Case "Cancel": CancelRoute run
        Case "Approve": ApproveSavedRoute run, Nz(f!txtNote, "")
        Case "Reopen": ReopenApprovedReport id, Nz(f!txtNote, "")
        Case Else: Err.Raise vbObjectError + 5903, "modUXReview.UXReviewAction", "Unknown review action."
    End Select
    Call UXRefreshEditorWorkflow(CLng(f!txtReportID))
    Forms!frmEditor!cmdClose.SetFocus
    f!txtNote = Null
    Call UXLoadReview
    Call UXWorkChanged
    UXReviewAction = True: Exit Function
Failed:
    UXReviewForm()!lblStatus.Caption = Err.Description
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
    UXGroupReview = UXOpenReportTool(id, "Review", Nz(f!cboGroup, 0))
End Function
'Send back goes to the earlier step selected in the steps list, else the previous
'step (step 0 is the requester).
Private Function UXSendBackTarget(ByVal runID As Long) As Long
    Dim f As Access.Form, k As Long
    Set f = UXReviewForm()
    k = Nz(DLookup("CurrentStep", "RouteRun", "RunID=" & runID), 0)
    If DLookup("RunState", "RouteRun", "RunID=" & runID) = "Accepted" Then k = DCount("*", "RouteRunStep", "RunID=" & runID)
    UXSendBackTarget = k - 1
    If Not IsNull(f!lstSteps) Then
        If CLng(f!lstSteps) < k Then UXSendBackTarget = CLng(f!lstSteps)
    End If
End Function
'Changes flagged for the holder since the report last moved: Members-grid edits
'and saves over a later reviewer's change.
Public Function RouteFlags(ByVal runID As Long) As String
    Dim r As DAO.Recordset, since As Variant
    since = DMax("RouteEventID", "RouteRunEvent", "RunID=" & runID & " AND EventKind In ('Start','Forward','Send back','Pull back')")
    If IsNull(since) Then Exit Function
    Set r = CurrentDb.OpenRecordset("SELECT e.Notes,e.EventAt,p.DisplayName FROM RouteRunEvent AS e LEFT JOIN UserProfile AS p ON e.AuthorProfileID=p.UserProfileID WHERE e.RunID=" & runID & " AND e.EventKind In ('Override','Grid change','Higher change') AND e.RouteEventID>" & since & " ORDER BY e.RouteEventID", dbOpenSnapshot)
    Do While Not r.EOF
        RouteFlags = RouteFlags & "Flagged: " & Nz(r!DisplayName, "Unknown contributor") & ", " & Format$(r!EventAt, "yyyy-mm-dd hh:nn") & ": " & Nz(r!Notes, "") & ". "
        r.MoveNext
    Loop
    r.Close
End Function
