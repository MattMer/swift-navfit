Attribute VB_Name = "UXNativeGroupsHarness"
Option Explicit
Private n As Long
Private rejection As String
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6942, , message
    n = n + 1
End Sub
Private Function RejectBatch(ByVal picks As Collection, ByVal hash As String) As Boolean
    Dim result As Collection
    On Error GoTo Rejected
    Set result = ImportNAVFITGroups(picks, hash, 2000): Exit Function
Rejected:
    rejection = Err.Description: RejectBatch = True
End Function
Public Function TestUXNativeGroups(ByVal source As String, ByVal bad As String, ByVal template As String, ByVal directory As String) As String
    Dim groups As Object, picks As New Collection, key As Variant, item As Object, hash As String, result As Collection, i As Long, id As Long
    Dim members As DAO.Recordset, native As DAO.Database, row As DAO.Recordset, saved As Object, decoded As Object, preview As Object, file As String
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC NATIVE GROUP IMPORT", "Importer"
    OpenReviewFile bad, False: Set groups = NAVFITReviewGroupChoices(2000)
    For Each key In groups.Keys: picks.Add CStr(key): hash = groups(key)("SourceHash"): Next
    Check RejectBatch(picks, hash), "Invalid last partition was accepted"
    Check InStr(rejection, "digits only") > 0, "Failure did not reach the last invalid partition: " & rejection
    Check DCount("*", "SummaryGroup") = 0 And DCount("*", "Report") = 0 And DCount("*", "ImportReceipt") = 0, "Failed batch left partial groups/reports/receipts"
    CloseReviewFile
    Check UXChooseNativeGroupFile(source), "Source chooser adapter failed"
    Check Forms!frmUXNativeGroups!lstGroups.ListCount = 3, "Initial discovery hid groups before century selection: " & Forms!frmUXNativeGroups!lblStatus.Caption
    Check UXLoadNativeGroups(), "Groups did not load: " & Forms!frmUXNativeGroups!lblStatus.Caption
    Check Forms!frmUXNativeGroups!lstGroups.ListCount = 3, "Native group partitions missing"
    Forms!frmUXNativeGroups!lstGroups.Selected(0) = True
    Check UXPreviewNativeGroups(), "Selected group preview failed: " & Forms!frmUXNativeGroups!lblStatus.Caption
    Check InStr(Forms!frmUXNativeGroups!txtPreview, "After:") > 0 And InStr(Forms!frmUXNativeGroups!txtPreview, "NEW GROUP:") > 0, "Full values/group details absent"
    Forms!frmUXNativeGroups!lstGroups.Selected(1) = True
    Check Not UXApplyNativeGroups(), "Changed selection accepted stale preview"
    Check DCount("*", "Report") = 0, "Rejected selection created report"
    Check UXPreviewNativeGroups(), "Updated selection preview failed"
    Check UXApplyNativeGroups(), "Selected groups apply failed: " & Forms!frmUXNativeGroups!lblStatus.Caption
    Check DCount("*", "SummaryGroup") = 2 And DCount("*", "Report") = 2 And DCount("*", "GroupMember") = 2, "Imported unselected groups or lost membership"
    Check DCount("*", "ImportReceipt", "AdapterVersion='navfit-full-v1/local-CStr'") = 2, "Group import omitted full provenance"
    Check Not UXPreviewNativeGroups(), "Duplicate source preview allowed"
    For i = 0 To 2: Forms!frmUXNativeGroups!lstGroups.Selected(i) = (i = 2): Next
    Check UXPreviewNativeGroups() And UXApplyNativeGroups(), "Remaining selected group failed"
    Check DCount("*", "SummaryGroup") = 3 And DCount("*", "Report") = 3, "Final selected import count"
    Check Not Forms!frmUXNativeGroups!cmdApply.Enabled, "Apply remained enabled after results"
    Check DCount("*", "FieldDefault", "ScopeKind='Group'") > 0, "Native shared folder defaults not imported"
    Set members = CurrentDb.OpenRecordset("SELECT GroupID,ReportID FROM GroupMember ORDER BY GroupID", dbOpenSnapshot)
    Do While Not members.EOF
        Set saved = SavedFields(members!ReportID): Set preview = NAVFITGroupPreview(members!GroupID)
        file = directory & "\imported-group-" & members!GroupID & ".accdb"
        NAVFITExportGroup members!GroupID, file, template, preview("Token")
        Set native = DBEngine.OpenDatabase(file, False, True)
        Set row = native.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
        Set decoded = NAVFITReadFields(row, 2000)
        For Each key In decoded.Keys: Check SameFieldValue(decoded(key), saved(key)), "Import/export round trip changed " & CStr(key): Next
        row.Close: native.Close: members.MoveNext
    Loop
    members.Close
    UXCloseNativeGroups
    TestUXNativeGroups = "PASS: " & n & " native group selection/atomic import assertions": Exit Function
Failed:
    TestUXNativeGroups = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
