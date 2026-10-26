Attribute VB_Name = "modUXHeaderReview"
Option Compare Binary
Option Explicit
Public Function UXHeaderReview(ByVal action As String, Optional ByVal decisionNote As Variant) As Boolean
 Dim f As Access.Form, choice As String, id As Long
 On Error GoTo Failed
 UXFlushSurface
 If Not IsMissing(decisionNote) And EditorHasPending() Then
  choice = UXAsk("Review")
  If choice = "Save" Then
   If Not UXSaveEditor() Then Exit Function
  ElseIf choice <> "Discard" Then
   Exit Function
  End If
 End If
 If Not UXToolsChoose("Review") Then Exit Function
 'Header buttons reflect the newest route; act on that route, not an older panel selection.
 id = Nz(Forms!frmEditor!txtReportID, 0)
 Call UXLoadReview(id, -1, Nz(MajorScalar("SELECT Max(RunID) FROM RouteRun WHERE ReportID=" & id), 0))
 Set f = UXReviewForm()
 If IsMissing(decisionNote) Then
  f!txtNote.SetFocus: UXHeaderReview = True: Exit Function
 End If
 If action = "Approval" Then
  If f!cmdReopen.Enabled Then action = "Reopen" Else action = "Approve"
 End If
 f!txtNote = CStr(decisionNote)
 UXHeaderReview = UXReviewAction(action)
 Exit Function
Failed:
 Forms!frmEditor!uxError.Caption = Err.Description
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
    Dim f As Access.Form, id As Long, run As Long, editable As Boolean, current As Boolean, author As Long, state As String, r As DAO.Recordset, reason As String, primary As String, nextAction As String, action As Variant, routeState As String
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
            run = r!RunID: current = RouteIsCurrent(run): routeState = r!RunState
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
        If state = "Approved" Then reason = "Locally approved: reopen approval with a reason before editing. Official submission is separate."
        If state = "Signed" Then reason = "Marked signed outside Swift: the first edit asks whether to reopen it for correction."
        If Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True) Then reason = "Recover this archived report before reviewing or reopening approval."
    End If
    If UXHasControl(f, "uxMarkSigned") Then f!uxMarkSigned.Enabled = id > 0 And (state = "Draft" Or state = "Approved") And Not Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True) And Not EditorHasPending()
    If id > 0 And Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True) Then
        nextAction = "Archived report. Recover it from Reports to continue."
    ElseIf state = "Signed" Then
        nextAction = SignedReportCaption(id)
    ElseIf id = 0 Or (EditorHasPending() And f!cmdSave.Enabled) Then
        primary = "cmdSave": nextAction = "Next: Save draft; review uses saved values."
    ElseIf f!uxReviewApproval.Enabled Then
        primary = "uxReviewApproval"
        If state = "Approved" Then nextAction = "Local approval recorded. Reopen to edit; this is not official submission." Else nextAction = "Next: Record local approval as the final reviewer."
    ElseIf f!uxReviewAccept.Enabled Then
        primary = "uxReviewAccept": nextAction = "Next: Accept this step or return the report with a decision note."
    ElseIf run > 0 And Not current And (routeState = "Pending" Or routeState = "Accepted") Then
        primary = "uxReviewSetup": nextAction = "Next: Open Review setup; this stale review needs the requester to restart it."
    ElseIf routeState = "Pending" Then
        primary = "uxReviewSetup": nextAction = "Waiting for the assigned reviewer. Review setup shows the current step."
    ElseIf routeState = "Accepted" Then
        primary = "uxReviewSetup": nextAction = "Waiting for the final reviewer to record local approval."
    Else
        primary = "uxReviewSetup": nextAction = "Next: Choose a route in Review setup. Local review is separate from submission."
    End If
    f!uxLocked.Caption = nextAction: f!uxLocked.ControlTipText = reason
    For Each action In Array("cmdSave", "uxReviewSetup", "uxReviewAccept", "uxReviewReturn", "uxReviewApproval")
        With f.Controls(action)
            .FontBold = (.Name = primary And .Enabled)
            ' Semantic colors are owned by UXApplyVisual below.
            If Not .Enabled Then .ControlTipText = "Unavailable: " & reason Else .ControlTipText = reason
        End With
    Next
    If Not f!uxReviewApproval.Enabled Then
        If routeState = "Accepted" And current Then f!uxReviewApproval.ControlTipText = "Unavailable: only the final reviewer can record local approval." Else f!uxReviewApproval.ControlTipText = Left$("Unavailable: complete all steps on a current route before local approval. " & reason, 255)
    End If
    If Not f!uxReviewAccept.Enabled And routeState <> "Pending" Then
        f!uxReviewAccept.ControlTipText = Left$("Unavailable: a current pending review step assigned to you is required. " & reason, 255)
        f!uxReviewReturn.ControlTipText = f!uxReviewAccept.ControlTipText
    End If
    UXApplyVisual f
    UXHeaderReviewState = True: Exit Function
Failed:
    On Error Resume Next
    If Not r Is Nothing Then r.Close
End Function
