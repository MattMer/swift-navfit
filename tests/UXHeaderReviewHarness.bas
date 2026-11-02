Attribute VB_Name = "UXHeaderReviewHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6911, , message
    n = n + 1
End Sub
Public Function TestUXHeaderReview() As String
    Dim id As Long, template As Long, run As Long, other As Long, c As Access.Control, name As Variant
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
    ' The retired review window is gone; header actions run through the report's embedded review panel.
    Check UXToolsView() = "Review" And UXToolsAvailable(), "Header action did not use the embedded review panel"
    Check CurrentProject.AllForms("frmEditor").IsLoaded, "Report disappeared"
    Check Forms!frmEditor!uxReviewApproval.Enabled, "Accepted route approval not enabled"
    Check Forms!frmEditor!uxReviewApproval.FontBold, "Final approval is not primary"
    Check UXHeaderReview("Approval", "Approved locally"), "Header local approval failed"
    Check DLookup("WorkflowState", "Report", "ReportID=" & id) = "Approved", "Approval not recorded"
    Check Forms!frmEditor!uxReviewApproval.Enabled And Forms!frmEditor!uxReviewApproval.Caption = "Reopen approval", "Reopen action missing"
    Check UXHeaderReview("Approval", "Reopened for correction"), "Header reopen failed"
    Check DLookup("WorkflowState", "Report", "ReportID=" & id) = "Draft", "Reopen not recorded"
    ' F01: Send back from step 1 returns the report to the requester; the route stays open.
    run = StartRoute(id, template)
    Check UXHeaderReview("Return", "Correct this report"), "Header send back failed"
    Check DLookup("RunState", "RouteRun", "RunID=" & run) = "Pending" And DLookup("CurrentStep", "RouteRun", "RunID=" & run) = 0, "Send back to the requester not recorded"
    ' The holder's own save keeps the route current, and Forward resubmits to step 1.
    Call CommitDraft(id, "EVAL", "Header review", "Synthetic member", #1/1/2026#, #9/30/2026#, "Holder change", "Test", "")
    Call UXHeaderReviewState
    Check Forms!frmEditor!uxReviewAccept.Enabled And Not Forms!frmEditor!uxReviewReturn.Enabled, "Requester cannot forward after editing"
    Check UXHeaderReview("Accept", "Resubmitted"), "Header forward failed"
    Check DLookup("RunState", "RouteRun", "RunID=" & run) = "Pending" And DLookup("CurrentStep", "RouteRun", "RunID=" & run) = 1, "Forward did not reach step 1"
    Check UXResizeEditor(9000, 7000), "Compact layout failed"
    For Each name In Array("uxReviewSetup", "uxReviewAccept", "uxReviewReturn", "uxReviewApproval")
        Set c = Forms!frmEditor.Controls(name)
        Check c.Visible And c.Left + c.Width < IIf(UXFixedEnabled(), Forms!frmEditor.Width, 9000), "Header clipped: " & c.Name
    Next
    TestUXHeaderReview = "PASS: " & n & " header review assertions": Exit Function
Failed:
    TestUXHeaderReview = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
