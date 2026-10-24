Attribute VB_Name = "MajorUIHarness"
Option Explicit
Public Function TestMajorUI() As String
    Dim f As Form, text As String, count As Long, first As Long, selected As Long, i As Long, id As Long
    On Error GoTo Failed
    'The retained adapter is hidden; Members is the canonical visible grid.
    DoCmd.OpenForm "frmMajorWorkbench", , , , , acHidden, "ContextOnly"
    If Not LoadMajorWorkbench() Then Err.Raise 5, , "Workbench adapter did not load"
    Set f = Forms!frmMajorWorkbench
    first = Abs(CInt(f!lstReports.ColumnHeads))
    If f!lstReports.ListCount - first <> 2 Then Err.Raise 5, , "Report selector did not load both data rows"
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
    selected = -1
    For i = first To f!lstReports.ListCount - 1
        id = CLng(f!lstReports.ItemData(i))
        If DLookup("WorkflowState", "Report", "ReportID=" & id) = "Draft" Then selected = i: Exit For
    Next
    If selected < 0 Then Err.Raise 5, , "Fixture has no editable report"
    f!lstReports.Selected(selected) = True: f!cboField = "OfficialComments": f!chkNull = False
    text = "Synthetic long value " & String$(5000, "x") & " END OF VALUE"
    f!txtValue = text: f!txtNote = "Synthetic UI preview"
    If Not MajorAction("PreviewBatch") Or Not MajorAction("ShowBatch") Then Err.Raise 5, , f!lblStatus.Caption
    first = Abs(CInt(f!lstDetails.ColumnHeads))
    If f!lstDetails.ListCount - first <> 1 Then Err.Raise 5, , "UI batch did not show staged item"
    count = count + 1
    f!lstDetails.Selected(first) = True
    If Not ShowMajorDetails() Then Err.Raise 5, , "Complete staged value unavailable"
    If InStr(Nz(f!txtDetail, ""), text) = 0 Then Err.Raise 5, , "Complete value truncated in review"
    count = count + 1
    TestMajorUI = "PASS: " & count & " retained workbench adapter assertions": Exit Function
Failed:
    TestMajorUI = "FAIL after " & count & ": " & Err.Number & " / " & Err.Description
End Function

Public Function TestMajorViewBounds() As String
    Dim name As Variant, f As Access.Form, c As Access.Control, count As Long, groupID As Long
    On Error GoTo Failed
    groupID = CLng(DMin("GroupID", "SummaryGroup"))
    If Not UXShowGroupMembers(groupID) Then Err.Raise 5, , "Current Members view did not open"
    For Each name In Array("frmUXComparison", "frmUXGroups", "frmPDFImport", "frmBackup")
        If name <> "frmUXComparison" Then If Not OpenScreen(CStr(name)) Then Err.Raise 5, , "Current destination did not open: " & name
        Set f = Forms(CStr(name))
        For Each c In f.Controls
            If c.Visible Then
                If c.Left < 0 Or c.Top < 0 Or c.Left + c.Width > f.Width Or c.Top + c.Height > f.Section(acDetail).Height Then Err.Raise 5, , "Control outside current form: " & name & "." & c.Name
                count = count + 1
            End If
        Next
    Next
    If count = 0 Then Err.Raise 5, , "No native controls inspected"
    TestMajorViewBounds = "PASS: " & count & " visible native controls within current view bounds": Exit Function
Failed:
    TestMajorViewBounds = "FAIL: " & Err.Number & " / " & Err.Description
End Function
