Attribute VB_Name = "MajorUIHarness"
Option Explicit
Public Function TestMajorUI() As String
    Dim f As Form, text As String, count As Long
    On Error GoTo Failed
    If Not OpenScreen("frmMajorWorkbench") Then Err.Raise 5, , "Workbench did not open"
    Set f = Forms!frmMajorWorkbench
    If f!lstReports.ListCount <> 2 Then Err.Raise 5, , "Report selector did not load"
    count = count + 1
    f!cboReviewer = MajorAuthor()
    If Not MajorAction("ClearReviewers") Then Err.Raise 5, , "Clear reviewer action failed"
    If Not MajorAction("AddReviewer") Or Not MajorAction("AddReviewer") Then Err.Raise 5, , "Reviewer selection failed"
    If f!txtReviewers <> MajorAuthor() & "," & MajorAuthor() Then Err.Raise 5, , "Ordered reviewer IDs lost"
    If InStr(f!txtReviewerNames, " > ") = 0 Then Err.Raise 5, , "Reviewer names not displayed"
    count = count + 2
    f!txtTitle = "Synthetic UI route": f!txtNote = "Synthetic UI instructions"
    If Not MajorAction("CreateRoute") Then Err.Raise 5, , f!lblStatus.Caption
    If DCount("*", "RouteTemplateStep", "TemplateID=" & f!cboTemplate) <> 2 Then Err.Raise 5, , "UI route steps missing"
    count = count + 1
    f!lstReports.Selected(0) = True: f!cboField = "OfficialComments": f!chkNull = False
    text = "Synthetic long value " & String$(5000, "x") & " END OF VALUE"
    f!txtValue = text: f!txtNote = "Synthetic UI preview"
    If Not MajorAction("PreviewBatch") Or Not MajorAction("ShowBatch") Then Err.Raise 5, , f!lblStatus.Caption
    If f!lstDetails.ListCount <> 1 Then Err.Raise 5, , "UI batch did not show staged item"
    count = count + 1
    f!lstDetails.SetFocus: f!lstDetails.Selected(0) = True
    If Not ShowMajorDetails() Then Err.Raise 5, , "Complete staged value unavailable"
    If InStr(Nz(f!txtDetail, ""), text) = 0 Then Err.Raise 5, , "Complete value truncated in review"
    count = count + 1
    TestMajorUI = "PASS: " & count & " workbench UI assertions": Exit Function
Failed:
    TestMajorUI = "FAIL after " & count & ": " & Err.Number & " / " & Err.Description
End Function
