Attribute VB_Name = "NativeGroupDiscoveryHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6950, , message
    n = n + 1
End Sub
Public Function TestOriginalGroupDiscovery(ByVal source As String) As String
    On Error GoTo Failed
    Check UXChooseNativeGroupFile(source), "Original source chooser"
    Check Forms!frmUXNativeGroups!lstGroups.ListCount = 3, "Original TEST_DB groups hidden"
    UXCloseNativeGroups
    TestOriginalGroupDiscovery = "PASS: original TEST_DB displays all three source groups": Exit Function
Failed:
    TestOriginalGroupDiscovery = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function TestNativeGroupDiscovery(ByVal source As String) As String
    Dim i As Long, j As Long, kind As String, grade As String, groups As Object, key As Variant, fields As Object
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC MISSING SOURCE RATE", "Importer"
    Check UXChooseNativeGroupFile(source), "Choose source"
    Check Forms!frmUXNativeGroups!lstGroups.ListCount = 3, "Missing rates hid source groups"
    Check UXSelectAllNativeGroups(True), "Select all failed"
    For i = 0 To 2: Check Forms!frmUXNativeGroups!lstGroups.Selected(i), "Select all omitted group": Next
    Forms!frmUXNativeGroups!cboImportGrade = "E5"
    Check Not UXAssignNativeGroupGrade(), "Mixed incompatible selection accepted"
    Set groups = NAVFITReviewGroupChoices(0, , True)
    For Each key In groups.Keys
        Check groups(key)("Paygrade") = "?", "Failed assignment mutated source metadata"
    Next
    Check Not UXPreviewNativeGroups(), "Missing grade allowed Preview"
    Check DCount("*", "Report") = 0, "Discovery created reports"
    For i = 0 To 2
        For j = 0 To 2: Forms!frmUXNativeGroups!lstGroups.Selected(j) = (i = j): Next
        kind = CStr(Forms!frmUXNativeGroups!lstGroups.Column(2, i))
        Select Case kind
            Case "EVAL": grade = "E5"
            Case "CHIEFEVAL": grade = "E7"
            Case "FITREP": grade = "O3"
        End Select
        Forms!frmUXNativeGroups!cboImportGrade = grade
        Check UXAssignNativeGroupGrade(), "Assign " & kind & ": " & Forms!frmUXNativeGroups!lblStatus.Caption
        Check Not Forms!frmUXNativeGroups!cmdApply.Enabled, "Assignment retained stale Apply"
        Check CStr(Forms!frmUXNativeGroups!lstGroups.Column(3, i)) = grade, "Assigned grade not visible"
    Next
    Check UXSelectAllNativeGroups(True), "Select all failed"
    For i = 0 To 2: Check Forms!frmUXNativeGroups!lstGroups.Selected(i), "Select all omitted group": Next
    Check UXPreviewNativeGroups(), "Preview failed: " & Forms!frmUXNativeGroups!lblStatus.Caption
    Check InStr(Forms!frmUXNativeGroups!txtPreview, "remain unchanged") > 0, "Missing source rate warning absent"
    Check UXSelectAllNativeGroups(False), "Clear selection failed"
    Check Not Forms!frmUXNativeGroups!cmdApply.Enabled, "Clear retained stale Apply"
    Check Not UXApplyNativeGroups(), "Cleared preview could still import"
    For i = 0 To 2: Check Not Forms!frmUXNativeGroups!lstGroups.Selected(i), "Clear omitted group": Next
    Check UXSelectAllNativeGroups(True), "Reselect all failed"
    Check Not Forms!frmUXNativeGroups!cmdApply.Enabled, "Select all retained stale Apply"
    Check UXPreviewNativeGroups(), "Reselected preview failed"
    Check UXApplyNativeGroups(), "Apply failed: " & Forms!frmUXNativeGroups!lblStatus.Caption
    Check DCount("*", "SummaryGroup") = 3 And DCount("*", "Report") = 3 And DCount("*", "GroupMember") = 3, "Imported group/member counts"
    Check DCount("*", "ImportReceipt") = 3, "Provenance missing"
    For Each key In Array(1, 2, 3)
        Set fields = SavedFields(CLng(key))
        Check Nz(fields("RankRate"), "") = "", "Import fabricated source rank/rate"
    Next
    Check Not UXPreviewNativeGroups(), "Duplicate import allowed"
    UXCloseNativeGroups
    TestNativeGroupDiscovery = "PASS: " & n & " missing-rate group discovery/import assertions": Exit Function
Failed:
    TestNativeGroupDiscovery = "FAIL: " & Err.Number & " " & Err.Description
End Function
