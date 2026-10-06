Attribute VB_Name = "modUXHeaderReview"
Option Compare Binary
Option Explicit
Public Function UXHeaderReview(ByVal action As String, Optional ByVal decisionNote As Variant) As Boolean
    Dim id As Long, choice As String, f As Access.Form, button As String, note As String, owned As Boolean
    On Error GoTo Failed
    UXFlushSurface
    If EditorHasPending() Then
        choice = UXAsk("Review")
        If choice = "Save" Then
            If Not UXSaveEditor() Then Exit Function
        ElseIf choice <> "Discard" Then
            Exit Function
        End If
    End If
    id = Nz(Forms!frmEditor!txtReportID, 0)
    If id = 0 Then Err.Raise vbObjectError + 6910, , "Save this new report before reviewing its saved revision."
    If CurrentProject.AllForms("frmUXReview").IsLoaded Then
        If Forms!frmUXReview.Visible Then Err.Raise vbObjectError + 6910, , "Finish the open review window first so its decision note is preserved."
        DoCmd.Close acForm, "frmUXReview", acSaveNo
    End If
    DoCmd.OpenForm "frmUXReview", , , , , acHidden, CStr(id): owned = True
    Set f = Forms!frmUXReview
    If action = "Approval" Then
        If f!cmdReopen.Enabled Then action = "Reopen" Else action = "Approve"
    End If
    button = "cmd" & action
    If Not f.Controls(button).Enabled Then Err.Raise vbObjectError + 6910, , f!lblStatus.Caption & " " & f!lblUnavailable.Caption
    If IsMissing(decisionNote) Then
        note = InputBox("Decision note for " & action & ". The action applies to the saved revision.", "Report review")
        If Len(Trim$(note)) = 0 Then GoTo Finished
    Else
        note = CStr(decisionNote)
    End If
    f!txtNote = note
    If Not UXReviewAction(action) Then Err.Raise vbObjectError + 6910, , f!lblStatus.Caption
    Forms!frmEditor!uxError.Caption = "Review updated: " & action & ". Decisions apply to saved report values."
    UXHeaderReview = True
Finished:
    If owned Then DoCmd.Close acForm, "frmUXReview", acSaveNo
    Exit Function
Failed:
    Forms!frmEditor!uxError.Caption = Err.Description
    Resume Finished
End Function
Public Function UXHeaderReviewLayout(ByVal compact As Boolean, ByVal width As Long) As Boolean
    Dim f As Access.Form, names As Variant, i As Long, x As Long, span As Long
    Set f = Forms!frmEditor
    names = Array("uxReviewSetup", "uxReviewAccept", "uxReviewReturn", "uxReviewApproval")
    If compact Then x = 250: span = width - 600 Else x = 5900: span = width - 6250
    For i = 0 To 3
        With f.Controls(names(i))
            .Visible = True: .Left = x + i * (span \ 4): .Top = 2650: .Width = span \ 4 - 80: .Height = 420
            If compact Then .Top = 2470
        End With
    Next
    Call UXHeaderReviewState
    UXHeaderReviewLayout = True
End Function
Public Function UXHeaderReviewState() As Boolean
    Dim f As Access.Form, id As Long, run As Long, editable As Boolean, current As Boolean, author As Long, state As String, r As DAO.Recordset, reason As String
    On Error GoTo Failed
    Set f = Forms!frmEditor: id = Nz(f!txtReportID, 0): author = MajorAuthor()
    f!uxReviewAccept.Enabled = False: f!uxReviewReturn.Enabled = False: f!uxReviewApproval.Enabled = False
    f!uxReviewSetup.Enabled = id > 0
    reason = "Save this new report before reviewing a saved revision."
    If id > 0 Then
        editable = Not Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True)
        state = Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "Draft")
        If state = "Approved" Then
            f!uxReviewApproval.Caption = "Reopen approval": f!uxReviewApproval.Enabled = editable
        Else
            f!uxReviewApproval.Caption = "Local approval"
        End If
        editable = editable And state = "Draft"
        reason = "Review setup starts a route. Decisions require the assigned contributor and a current saved revision; local approval requires the final reviewer."
        Set r = CurrentDb.OpenRecordset("SELECT TOP 1 * FROM RouteRun WHERE ReportID=" & id & " ORDER BY RunID DESC", dbOpenSnapshot)
        If Not r.EOF Then
            run = r!RunID: current = RouteIsCurrent(run)
            If Not current And (r!RunState = "Pending" Or r!RunState = "Accepted") Then
                reason = "Stale review: open Review setup and restart the route against current saved values. Only its requester can restart it."
            ElseIf r!RunState = "Pending" Then
                f!uxReviewAccept.Enabled = editable And current And Nz(DLookup("AssigneeID", "RouteRunStep", "RunID=" & run & " AND StepOrdinal=" & r!CurrentStep), 0) = author
                f!uxReviewReturn.Enabled = f!uxReviewAccept.Enabled
                reason = "Current step assigned to " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & Nz(DLookup("AssigneeID", "RouteRunStep", "RunID=" & run & " AND StepOrdinal=" & r!CurrentStep), 0)), "Unknown contributor") & ". Decisions apply to saved values."
            ElseIf r!RunState = "Accepted" And state <> "Approved" Then
                f!uxReviewApproval.Enabled = editable And current And Nz(r!ApprovedBy, 0) = author
                reason = "All steps accepted. Local approval requires the final reviewer and a decision note."
            End If
        End If
        r.Close
        If Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True) Then reason = "Recover this archived report before reviewing or reopening approval."
    End If
    f!uxReviewAccept.ControlTipText = reason: f!uxReviewReturn.ControlTipText = reason: f!uxReviewApproval.ControlTipText = reason
    UXHeaderReviewState = True: Exit Function
Failed:
    On Error Resume Next
    If Not r Is Nothing Then r.Close
End Function
