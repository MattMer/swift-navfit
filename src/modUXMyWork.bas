Attribute VB_Name = "modUXMyWork"
Option Compare Binary
Option Explicit

Public Function UXWorkReport(ByVal id As Long) As String
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    If Not r.EOF Then UXWorkReport = UXIdentity(r!OfficialName, r!MemberName) & " | " & r!FormKind & " | " & Nz(Format(r!PeriodStart, "yyyy-mm-dd"), "No start date") & " to " & Nz(Format(r!PeriodEnd, "yyyy-mm-dd"), "No end date")
    r.Close
End Function
Public Function UXThreadOpen(ByVal id As Long) As Boolean
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT TOP 1 EventKind FROM ReviewCommentEvent WHERE ThreadID=" & id & " ORDER BY CommentEventID DESC", dbOpenSnapshot)
    UXThreadOpen = True
    If Not r.EOF Then UXThreadOpen = (r!EventKind <> "Resolved")
    r.Close
End Function
Private Function WorkRecord(ByVal kind As String, ByVal id As Long) As DAO.Recordset
    Select Case kind
        Case "Assignment": Set WorkRecord = CurrentDb.OpenRecordset("SELECT * FROM ReviewTask WHERE TaskID=" & id, dbOpenSnapshot)
        Case "Route": Set WorkRecord = CurrentDb.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & id, dbOpenSnapshot)
        Case "Comment": Set WorkRecord = CurrentDb.OpenRecordset("SELECT * FROM ReviewComment WHERE CommentID=" & id & " AND ThreadID Is Null", dbOpenSnapshot)
    End Select
End Function
Public Function UXWorkVisible(ByVal kind As String, ByVal id As Long, ByVal pid As Long) As Boolean
    Dim r As DAO.Recordset, report As Long, assignee As Long
    Set r = WorkRecord(kind, id)
    If r.EOF Then r.Close: Exit Function
    report = r!ReportID
    Select Case kind
        Case "Assignment"
            UXWorkVisible = r!TaskState = "Pending" And (r!AssignedTo = pid Or (r!RequestedBy = pid And r!RevisionID <> CurrentRevision(report)))
        Case "Route"
            If r!RunState = "Pending" Then
                assignee = Nz(DLookup("AssigneeID", "RouteRunStep", "RunID=" & id & " AND StepOrdinal=" & r!CurrentStep), 0)
                UXWorkVisible = assignee = pid Or (r!RequestedBy = pid And Not RouteIsCurrent(id))
            ElseIf r!RunState = "Accepted" Then
                UXWorkVisible = Nz(r!ApprovedBy, 0) = pid Or (r!RequestedBy = pid And Not RouteIsCurrent(id))
            End If
        Case "Comment"
            If UXThreadOpen(id) Then
                UXWorkVisible = DCount("*", "ReviewComment", "(CommentID=" & id & " OR ThreadID=" & id & ") AND AuthorProfileID=" & pid) > 0
                If Not UXWorkVisible Then UXWorkVisible = DCount("*", "ReportRevision", "ReportID=" & report & " AND AuthorProfileID=" & pid) > 0
                If Not UXWorkVisible Then UXWorkVisible = DCount("*", "ReviewTask", "ReportID=" & report & " AND TaskState='Pending' AND (AssignedTo=" & pid & " OR RequestedBy=" & pid & ")") > 0
                If Not UXWorkVisible Then UXWorkVisible = DCount("*", "RouteRun", "ReportID=" & report & " AND RunState In ('Pending','Accepted') AND (RequestedBy=" & pid & " OR ApprovedBy=" & pid & ")") > 0
                If Not UXWorkVisible Then UXWorkVisible = Nz(MajorScalar("SELECT Count(*) FROM RouteRunStep AS s INNER JOIN RouteRun AS u ON s.RunID=u.RunID WHERE u.ReportID=" & report & " AND u.RunState='Pending' AND s.StepOrdinal=u.CurrentStep AND s.AssigneeID=" & pid), 0) > 0
            End If
    End Select
    r.Close
End Function
Public Function UXWorkText(ByVal kind As String, ByVal id As Long, ByVal part As String) As String
    Dim r As DAO.Recordset, status As String, action As String, owner As Long, label As String, stale As Boolean, report As Long
    Set r = WorkRecord(kind, id)
    If r.EOF Then r.Close: Exit Function
    report = r!ReportID
    Select Case kind
        Case "Assignment"
            label = "Assigned review": owner = r!AssignedTo: status = r!TaskState
            If status = "Pending" Then
                stale = r!RevisionID <> CurrentRevision(report)
                If stale Then status = "Stale": action = "Requester: cancel and reassign" Else action = "Open report; accept or return"
                If Not IsNull(r!DueDate) Then
                    If r!DueDate < Date Then status = status & " / overdue"
                End If
            Else
                action = "View review history"
            End If
            If part = "Due" Then UXWorkText = Nz(Format(r!DueDate, "yyyy-mm-dd"), "Not set")
        Case "Route"
            status = r!RunState: label = "Route: step " & r!CurrentStep & " of " & DCount("*", "RouteRunStep", "RunID=" & id)
            owner = Nz(DLookup("AssigneeID", "RouteRunStep", "RunID=" & id & " AND StepOrdinal=" & r!CurrentStep), 0)
            If status = "Pending" Or status = "Accepted" Then
                stale = Not RouteIsCurrent(id)
                If stale Then
                    status = "Stale": action = "Requester: restart route": owner = r!RequestedBy
                ElseIf status = "Accepted" Then
                    label = "Route: all steps accepted": action = "Final reviewer: local approval": owner = Nz(r!ApprovedBy, 0)
                Else
                    action = "Assigned reviewer: accept or return"
                End If
            Else
                action = "View route history"
            End If
            If part = "Due" Then UXWorkText = "Not set"
        Case "Comment"
            label = "Comment thread": owner = r!AuthorProfileID
            If UXThreadOpen(id) Then status = "Unresolved": action = "Read, reply or resolve" Else status = "Resolved": action = "View thread history"
            If r!RevisionID <> CurrentRevision(report) Then status = status & " / earlier revision"
            If part = "Due" Then UXWorkText = "Not set"
    End Select
    If Nz(DLookup("IsArchived", "Report", "ReportID=" & report), True) Then
        status = status & " / archived": action = "Recover report before review activity"
    ElseIf Nz(DLookup("WorkflowState", "Report", "ReportID=" & report), "") = "Approved" And kind <> "Comment" Then
        status = status & " / local approval": action = "Reopen local approval before decisions"
    End If
    Select Case part
        Case "Label": UXWorkText = label
        Case "Owner": UXWorkText = Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & owner), "Unknown contributor")
        Case "Status": UXWorkText = status
        Case "Action": UXWorkText = action
    End Select
    r.Close
End Function
Public Function UXWorkSQL(ByVal pid As Long, Optional ByVal history As Boolean = False) As String
    Dim kind As Variant, table As String, key As String, sql As String, branch As String
    For Each kind In Array("Assignment", "Route", "Comment")
        Select Case kind
            Case "Assignment": table = "ReviewTask": key = "TaskID"
            Case "Route": table = "RouteRun": key = "RunID"
            Case "Comment": table = "ReviewComment": key = "CommentID"
        End Select
        branch = "SELECT '" & kind & "' & ':' & " & key & " AS WorkKey,'" & kind & "' AS WorkKind," & key & " AS ItemID,ReportID,RevisionID,UXWorkReport(ReportID) AS [Member and report],UXWorkText('" & kind & "'," & key & ",'Label') AS [Work item],UXWorkText('" & kind & "'," & key & ",'Owner') AS [Contributor],UXWorkText('" & kind & "'," & key & ",'Due') AS [Due date],UXWorkText('" & kind & "'," & key & ",'Status') AS [Status],UXWorkText('" & kind & "'," & key & ",'Action') AS [Next action] FROM " & table & " WHERE "
        If kind = "Comment" Then branch = branch & "ThreadID Is Null AND "
        If history Then branch = branch & "True" Else branch = branch & "UXWorkVisible('" & kind & "'," & key & "," & pid & ")"
        If Len(sql) > 0 Then sql = sql & " UNION ALL "
        sql = sql & branch
    Next
    UXWorkSQL = sql & " ORDER BY [Member and report],WorkKind,ItemID"
End Function
Public Function UXWorkCount() As Long
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset(UXWorkSQL(MajorAuthor()), dbOpenSnapshot)
    If Not r.EOF Then r.MoveLast: UXWorkCount = r.RecordCount
    r.Close
End Function
Public Function UXLoadWorkQueue() As Boolean
    Dim f As Access.Form, selected As String, i As Long
    If Not CurrentProject.AllForms("frmMyWork").IsLoaded Then Exit Function
    Set f = Forms!frmMyWork
    If Not UXHasControl(f, "lstWork") Then Exit Function
    If UXHasControl(f, "uxContributor") Then
        f!uxContributor.RowSource = "SELECT UserProfileID,DisplayName FROM UserProfile ORDER BY DisplayName"
        f!uxContributor = MajorAuthor()
    End If
    selected = Nz(f!lstWork, "")
    f!lstWork.RowSource = UXWorkSQL(MajorAuthor(), Nz(f!chkAll, False))
    f!lstWork = Null
    For i = 1 To f!lstWork.ListCount - 1
        If Nz(f!lstWork.Column(0, i), "") = selected Then f!lstWork = selected
    Next
    f!uxWorkStatus.Caption = IIf(f!lstWork.ListCount <= 1, "No work items in this view.", (f!lstWork.ListCount - 1) & " work items. Select a row to see its next action.")
    Call UXWorkSelection
    UXLoadWorkQueue = True
End Function
Public Function UXWorkSelection() As Boolean
    Dim f As Access.Form, kind As String, id As Long, r As DAO.Recordset, pid As Long, canDecide As Boolean, text As String
    Set f = Forms!frmMyWork: pid = MajorAuthor()
    f!cmdAccept.Enabled = False: f!cmdReturn.Enabled = False: f!cmdCancel.Enabled = False
    f!cmdDraft.Enabled = Not IsNull(f!lstWork): f!cmdOpenWork.Enabled = Not IsNull(f!lstWork)
    f!txtHistory = Null
    If IsNull(f!lstWork) Then
        If f!lstWork.ListCount <= 1 Then f!uxWorkStatus.Caption = "No work items in this view." Else f!uxWorkStatus.Caption = (f!lstWork.ListCount - 1) & " work items. Select a row to see its next action."
        Exit Function
    End If
    kind = f!lstWork.Column(1): id = f!lstWork.Column(2)
    Set r = WorkRecord(kind, id)
    If r.EOF Then r.Close: Exit Function
    text = UXWorkReport(r!ReportID) & vbCrLf & "Saved revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & r!RevisionID), "?") & " | " & UXWorkText(kind, id, "Status") & vbCrLf & "Next action: " & UXWorkText(kind, id, "Action")
    If kind = "Assignment" Then
        f!lstTasks = id
        canDecide = r!TaskState = "Pending" And r!AssignedTo = pid And r!RevisionID = CurrentRevision(r!ReportID) And Not Nz(DLookup("IsArchived", "Report", "ReportID=" & r!ReportID), True) And Nz(DLookup("WorkflowState", "Report", "ReportID=" & r!ReportID), "") = "Draft"
        f!cmdAccept.Enabled = canDecide: f!cmdReturn.Enabled = canDecide
        f!cmdCancel.Enabled = r!TaskState = "Pending" And r!RequestedBy = pid
        text = text & vbCrLf & "Requested by " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & r!RequestedBy), "") & vbCrLf & Nz(r!Instructions, "")
    ElseIf kind = "Comment" Then
        text = text & vbCrLf & "Related unresolved threads are included for authors and active reviewers; comments are not assigned tasks." & vbCrLf & Nz(r!CommentBody, "")
    Else
        text = text & vbCrLf & UXRouteLabel(id)
    End If
    r.Close: f!txtHistory = text
    f!uxWorkStatus.Caption = "Assignment decisions use saved revisions. Routes and comments open their focused review screens."
    UXWorkSelection = True
End Function
Public Function UXWorkOpen(Optional ByVal reportOnly As Boolean = False) As Boolean
    Dim f As Access.Form, id As Long, report As Long, kind As String, decision As String, group As Long
    On Error GoTo Failed
    Set f = Forms!frmMyWork
    If IsNull(f!lstWork) Then Exit Function
    kind = f!lstWork.Column(1): id = f!lstWork.Column(2): report = f!lstWork.Column(3)
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Nz(Forms!frmEditor!txtReportID, 0) <> report Or reportOnly Then
            If Not UXLeaveEditor() Then Exit Function
        ElseIf EditorHasPending() Then
            decision = UXAsk("Review")
            If decision = "Save" Then
                If Not UXSaveEditor() Then Exit Function
            ElseIf decision <> "Discard" Then
                Exit Function
            End If
        End If
    End If
    If reportOnly Or kind = "Assignment" Then
        DoCmd.OpenForm "frmEditor", , , , , , CStr(report)
    ElseIf kind = "Route" Then
        group = Nz(DLookup("GroupID", "RouteRun", "RunID=" & id), 0)
        If CurrentProject.AllForms("frmUXReview").IsLoaded Then DoCmd.Close acForm, "frmUXReview", acSaveNo
        DoCmd.OpenForm "frmUXReview", , , , , , report & ":" & group & ":" & id
    Else
        Call OpenReviewComments(report)
        Forms!frmReviewComments!lstThreads = id: Call ShowReviewThread
    End If
    UXWorkOpen = True: Exit Function
Failed:
    Forms!frmMyWork!uxWorkStatus.Caption = Err.Description
End Function
Public Function UXWorkChanged() As Boolean
    If CurrentProject.AllForms("frmMyWork").IsLoaded Then Call UXLoadWorkQueue
    Call UXDashboardRefresh
    UXWorkChanged = True
End Function
Public Function UXWorkBack() As Boolean
    If Not UXLeaveEditor() Then Exit Function
    DoCmd.OpenForm "frmMyWork": Call LoadMyWork: Call UXMyWorkMode("Assignments")
    UXWorkBack = True
End Function
Public Function UXWorkLayout(Optional ByVal width As Long = 0, Optional ByVal height As Long = 0) As Boolean
    Dim f As Access.Form, size As Long, parts As Variant, part As Variant, columns As String
    If Not CurrentProject.AllForms("frmMyWork").IsLoaded Then Exit Function
    Set f = Forms!frmMyWork
    If Not UXHasControl(f, "lstWork") Then Exit Function
    Call UXResizeWorkspace("frmMyWork", width, height)
    If UXHasControl(f, "uxContributor") And IIf(width > 0, width, f.InsideWidth) < 12000 Then
        f!uxContributorLabel.Left = f!lstWork.Left: f!uxContributorLabel.Top = f!lstWork.Top - 1550
        f!uxContributorLabel.Width = f!lstWork.Width
        f!uxContributor.Left = f!lstWork.Left: f!uxContributor.Top = f!lstWork.Top - 1200
        f!uxContributor.Width = f!lstWork.Width
        f!chkAll.Top = f!lstWork.Top - 700: f!lblAll.Top = f!chkAll.Top
        f!uxWorkStatus.Top = f!lstWork.Top - 350: f!uxWorkStatus.Height = 300
    End If
    size = f!lstWork.Width - 350: columns = "0;0;0;0;0"
    For Each part In Array(0.28, 0.15, 0.16, 0.1, 0.14, 0.17)
        columns = columns & ";" & CLng(size * CDbl(part))
    Next
    f!lstWork.ColumnWidths = columns
    UXWorkLayout = True
End Function
