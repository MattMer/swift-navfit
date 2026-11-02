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
        Case "Comment"
            If id < 0 Then
                Set WorkRecord = CurrentDb.OpenRecordset("SELECT ReportID,ContextRevisionID AS RevisionID,AuthorProfileID FROM ReportMessage WHERE MessageID=" & -id & " AND ParentMessageID Is Null", dbOpenSnapshot)
            Else
                Set WorkRecord = CurrentDb.OpenRecordset("SELECT * FROM ReviewComment WHERE CommentID=" & id & " AND ThreadID Is Null", dbOpenSnapshot)
            End If
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
        Case "Comment": UXWorkVisible = False
    End Select
    r.Close
    ' Review work on a signed report waits until it is reopened.
    If UXWorkVisible Then UXWorkVisible = Nz(DLookup("WorkflowState", "Report", "ReportID=" & report), "") <> "Signed"
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
            status = "Retained history": action = "Open the report conversation in Notes"
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
Public Function UXWorkSQL(ByVal pid As Long, Optional ByVal history As Boolean = False, Optional ByVal prepared As Boolean = False) As String
    Dim kind As Variant, table As String, key As String, sql As String, branch As String
    For Each kind In Array("Assignment", "Route")
        Select Case kind
            Case "Assignment": table = "ReviewTask": key = "TaskID"
            Case "Route": table = "RouteRun": key = "RunID"
            Case "Comment": table = "ReviewComment": key = "CommentID"
        End Select
        branch = "SELECT '" & kind & "' & ':' & " & key & " AS WorkKey,'" & kind & "' AS WorkKind," & key & " AS ItemID,ReportID,RevisionID,UXWorkReport(ReportID) AS [Member and report],UXWorkText('" & kind & "'," & key & ",'Label') AS [Work item],UXWorkText('" & kind & "'," & key & ",'Owner') AS [Contributor],UXWorkText('" & kind & "'," & key & ",'Due') AS [Due date],UXWorkText('" & kind & "'," & key & ",'Status') AS [Status],UXWorkText('" & kind & "'," & key & ",'Action') AS [Next action] FROM " & table & " WHERE "
        If prepared Then
            branch = "SELECT '" & kind & "' & ':' & " & key & " AS WorkKey,'" & kind & "' AS WorkKind," & key & " AS ItemID,ReportID,RevisionID,UXWorkRowCell('" & kind & "'," & key & ",0) AS [Member and report],UXWorkRowCell('" & kind & "'," & key & ",1) AS [Work item],UXWorkRowCell('" & kind & "'," & key & ",2) AS [Contributor],UXWorkRowCell('" & kind & "'," & key & ",3) AS [Due date],UXWorkRowCell('" & kind & "'," & key & ",4) AS [Status],UXWorkRowCell('" & kind & "'," & key & ",5) AS [Next action] FROM " & table & " WHERE UXWorkRowExists('" & kind & "'," & key & ")"
        Else
            branch = branch & WorkFilter(CStr(kind), key, pid, history)
        End If
        If Len(sql) > 0 Then sql = sql & " UNION ALL "
        sql = sql & branch
    Next
    UXWorkSQL = sql & " ORDER BY [Member and report],WorkKind,ItemID"
End Function
Private Function WorkFilter(ByVal kind As String, ByVal key As String, ByVal pid As Long, ByVal history As Boolean) As String
    If kind = "Comment" Then WorkFilter = "ThreadID Is Null AND "
    If history Then
        WorkFilter = WorkFilter & "True"
    Else
        WorkFilter = WorkFilter & "UXWorkVisible('" & kind & "'," & key & "," & pid & ")"
    End If
End Function
Public Function UXWorkCount() As Long
    Dim visible As Object
    ' A fresh inclusion snapshot only; do not materialize presentation or replace it.
    Set visible = UXWorkVisibleRows(MajorAuthor())
    UXWorkCount = visible.Count
End Function
Public Function UXLoadWorkQueue() As Boolean
    Dim f As Access.Form, selected As String, i As Long
    If Not CurrentProject.AllForms("frmMyWork").IsLoaded Then Exit Function
    Call UXAssertView("frmMyWork")
    Set f = Forms!frmMyWork
    f!uxContributor.RowSource = "SELECT UserProfileID,DisplayName FROM UserProfile ORDER BY DisplayName"
    f!uxContributor = MajorAuthor()
    selected = Nz(f!lstWork, "")
    Call UXPrepareWorkRows(MajorAuthor(), Nz(f!chkAll, False))
    f!lstWork.RowSource = UXWorkSQL(MajorAuthor(), Nz(f!chkAll, False), True)
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
    f!uxWorkStatus.Caption = "Assignment decisions use saved revisions. Open selected work opens the report Review steps view. Shared communication is in the Report notes tab."
    UXWorkSelection = True
End Function
Public Function UXWorkOpen(Optional ByVal reportOnly As Boolean = False) As Boolean
 Dim f As Access.Form, id As Long, report As Long, kind As String, group As Long, message As Long
 On Error GoTo Failed
 Set f = Forms!frmMyWork
 If IsNull(f!lstWork) Then Exit Function
 kind = f!lstWork.Column(1): id = f!lstWork.Column(2): report = f!lstWork.Column(3)
 If reportOnly Then
  If Not UXOpenReportTool(report, "Notes") Then Exit Function
  UXWorkOpen = UXSelectSection(0)
 ElseIf kind = "Assignment" Then
  UXWorkOpen = UXOpenReportTool(report, "Review", 0, 0, id)
 ElseIf kind = "Route" Then
  group = Nz(DLookup("GroupID", "RouteRun", "RunID=" & id), 0)
  UXWorkOpen = UXOpenReportTool(report, "Review", group, id)
 Else
  MigrateReportMessages report
  message = Nz(DLookup("MessageID", "ReportMessage", "LegacyCommentID=" & id), 0)
  UXWorkOpen = UXOpenReportTool(report, "Notes", 0, 0, 0, message)
 End If
 Exit Function
Failed:
 Forms!frmMyWork!uxWorkStatus.Caption = Err.Description
End Function
Public Function UXWorkChanged() As Boolean
    If CurrentProject.AllForms("frmMyWork").IsLoaded Then Call UXLoadWorkQueue
    Call UXDashboardRefresh
    UXWorkChanged = True
End Function
Public Function UXWorkBack() As Boolean
 Dim wasLoaded As Boolean
 If Not UXLeaveEditor() Then Exit Function
 wasLoaded = CurrentProject.AllForms("frmMyWork").IsLoaded
 DoCmd.OpenForm "frmMyWork"
 If wasLoaded Then UXLoadMyWork
 Forms!frmMyWork.SetFocus
 UXWorkBack = True
End Function
'Counseling tab: members overdue, due this month, due within 60 days, needing a
'rank/rate, or with a worksheet disposition due after detachment.
Public Function UXMyWorkCounseling() As Boolean
    Dim f As Access.Form, rows As String
    If Not CurrentProject.AllForms("frmMyWork").IsLoaded Then Exit Function
    Set f = Forms!frmMyWork
    If Not UXHasControl(f, "lstCounseling") Then Exit Function
    rows = CounselingWorkRows(Nz(f!chkCounselingAll, False))
    f!lstCounseling.RowSourceType = "Value List": f!lstCounseling.RowSource = rows
    f!lblCounselingScope.Caption = IIf(rows = "", "No counseling needs attention. ", "") & "Midterm counseling follows the BUPERSINST 1610.10H chapter 18 schedule by paygrade. Select a member to record a session or open their history."
    UXMyWorkCounseling = True
End Function
Private Function UXMyWorkCounselingMemberID() As Long
    If IsNull(Forms!frmMyWork!lstCounseling) Then Forms!frmMyWork!lblCounselingScope.Caption = "Select a member first.": Exit Function
    UXMyWorkCounselingMemberID = CLng(Forms!frmMyWork!lstCounseling)
End Function
Public Function UXMyWorkCounselingRecord() As Boolean
    Dim member As Long
    member = UXMyWorkCounselingMemberID(): If member = 0 Then Exit Function
    UXMyWorkCounselingRecord = UXCounselingOpen(True, member)
End Function
Public Function UXMyWorkCounselingMember() As Boolean
    Dim member As Long
    member = UXMyWorkCounselingMemberID(): If member = 0 Then Exit Function
    UXMyWorkCounselingMember = UXDirectoryHistoryFor(member)
End Function