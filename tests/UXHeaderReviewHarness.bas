Attribute VB_Name = "UXHeaderReviewHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6911, , message
    n = n + 1
End Sub
Public Function TestUXHeaderReview() As String
    Dim id As Long, template As Long, run As Long, other As Long, c As Access.Control
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC HEADER REVIEW", "Reviewer"
    id = CommitDraft(0, "EVAL", "Header review", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test", "")
    template = CreateRouteTemplate("Header route", CStr(MajorAuthor()), "Check saved values")
    run = StartRoute(id, template)
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Check Forms!frmEditor!uxReviewAccept.Enabled And Not Forms!frmEditor!uxReviewApproval.Enabled, "Initial review availability"
    Check Forms!frmEditor!uxReviewAccept.FontBold, "Assigned accept is not primary"
    Check UXHeaderReview("Accept", "Step checked"), "Header accept failed: " & Forms!frmEditor!uxError.Caption
    Check DLookup("RunState", "RouteRun", "RunID=" & run) = "Accepted", "Accept not recorded"
    Check Not CurrentProject.AllForms("frmUXReview").IsLoaded, "Hidden review adapter remained open"
    Check CurrentProject.AllForms("frmEditor").IsLoaded, "Report disappeared"
    Check Forms!frmEditor!uxReviewApproval.Enabled, "Accepted route approval not enabled"
    Check Forms!frmEditor!uxReviewApproval.FontBold, "Final approval is not primary"
    Check UXHeaderReview("Approval", "Approved locally"), "Header local approval failed"
    Check DLookup("WorkflowState", "Report", "ReportID=" & id) = "Approved", "Approval not recorded"
    Check Forms!frmEditor!uxReviewApproval.Enabled And Forms!frmEditor!uxReviewApproval.Caption = "Reopen approval", "Reopen action missing"
    Check UXHeaderReview("Approval", "Reopened for correction"), "Header reopen failed"
    Check DLookup("WorkflowState", "Report", "ReportID=" & id) = "Draft", "Reopen not recorded"
    run = StartRoute(id, template)
    Check UXHeaderReview("Return", "Correct this report"), "Header return failed"
    Check DLookup("RunState", "RouteRun", "RunID=" & run) = "Returned", "Return not recorded"
    run = StartRoute(id, template)
    Call CommitDraft(id, "EVAL", "Header review", "Synthetic member", #1/1/2026#, #9/30/2026#, "Concurrent change", "Test", "")
    Call UXHeaderReviewState
    Check Not Forms!frmEditor!uxReviewAccept.Enabled And InStr(Forms!frmEditor!uxReviewAccept.ControlTipText, "Stale review") > 0, "Stale action not visibly disabled"
    Check Not UXHeaderReview("Accept", "Cannot accept stale values"), "Stale route accepted"
    Check InStr(Forms!frmEditor!uxError.Caption, "Needs restart") > 0, "Stale next action missing"
    Check DLookup("RunState", "RouteRun", "RunID=" & run) = "Pending", "Rejected action changed route"
    Check UXResizeEditor(9000, 7000), "Compact layout failed"
    For Each c In Forms!frmEditor.Controls
        If Left$(c.Name, 8) = "uxReview" Then Check c.Visible And c.Left + c.Width < IIf(UXFixedEnabled(), Forms!frmEditor.Width, 9000), "Header clipped: " & c.Name
    Next
    TestUXHeaderReview = "PASS: " & n & " header review assertions": Exit Function
Failed:
    TestUXHeaderReview = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
