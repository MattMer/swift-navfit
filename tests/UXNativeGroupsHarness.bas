Attribute VB_Name = "UXNativeGroupsHarness"
Option Explicit
Private n As Long
Private rejection As String
Private batchOutcome As String
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6942, , message
    n = n + 1
End Sub
Private Function RejectBatch(ByVal picks As Collection, ByVal hash As String) As Boolean
    Dim result As Collection
    On Error GoTo Rejected
    Set result = ImportNAVFITGroups(picks, hash, 2000, outcome:=batchOutcome): Exit Function
Rejected:
    rejection = Err.Description: RejectBatch = True
End Function
Public Function TestNativeGroupApplyFault(ByVal source As String, ByVal fault As String) As String
    Dim before As Long, i As Long, message As String, context As Object
    On Error GoTo Failed
    n = 0: ConfigureWorkspace "SYNTHETIC GROUP APPLY FAULT", "Synthetic importer"
    Check UXChooseNativeGroupFile(source), "Fault source did not open"
    For i = 0 To Forms!frmUXNativeGroups!lstGroups.ListCount - 1
        Forms!frmUXNativeGroups!lstGroups.Selected(i) = (fault = "Rollback" Or i = 0)
    Next
    Check UXPreviewNativeGroups(), "Fault preview failed"
    Check Not UXApplyNativeGroups(), "Injected apply failure was accepted"
    message = Forms!frmUXNativeGroups!lblStatus.Caption
    Check Not Forms!frmUXNativeGroups!cmdApply.Enabled, "Failed Apply remains enabled"
    before = IIf(fault = "Committed" Or fault = "Display", 1, 0)
    Check DCount("*", "Report") = before And DCount("*", "SummaryGroup") = before And DCount("*", "ImportReceipt") = before, "Apply outcome disagrees with saved data"
    Select Case fault
        Case "Rollback": Check InStr(message, "Import rolled back") > 0, "Confirmed rollback not reported"
        Case "Uncertain": Check InStr(message, "could not be confirmed") > 0 And InStr(message, "no selected-group") = 0, "Uncertain outcome claims rollback"
        Case Else: Check InStr(message, "Import committed") > 0 And InStr(message, "rolled back") = 0, "Committed import claims rollback"
    End Select
    If fault = "Committed" Then
        Set context = SwiftLastFailureContext()
        Check context("Number") = vbObjectError + 9691 And context("Source") = "Synthetic native group apply", "Postcommit primary identity lost"
        Check context("Transaction") = "Committed", "Postcommit diagnostic lost transaction outcome"
    End If
    Check Not UXApplyNativeGroups(), "Failed Apply retained actionable preview"
    Check DCount("*", "Report") = before And DCount("*", "ImportReceipt") = before, "Repeated Apply duplicated data"
    Check CloseReviewFile(), "Fault source reader did not close"
    TestNativeGroupApplyFault = "PASS: " & n & " native group " & fault & " outcome/duplicate prevention assertions": Exit Function
Failed:
    TestNativeGroupApplyFault = "FAIL: " & fault & ": " & Err.Number & " " & Err.Description
End Function
Public Function PrepareOwnedPreviewReset(ByVal source As String) As String
    Dim report As Long, group As Long, fields As Object, first As Long, second As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC PREVIEW RESET", "Synthetic reviewer"
    If Not UXChooseNativeGroupFile(source) Then Err.Raise 5, , "Reset source did not open"
    Forms!frmUXNativeGroups!lstGroups.Selected(0) = True
    If Not UXPreviewNativeGroups() Then Err.Raise 5, , "Reset import preview failed"
    report = CommitDraft(0, "EVAL", "Synthetic move reset", "SYNTHETIC RESET MEMBER", Null, Null, "")
    Set fields = CreateObject("Scripting.Dictionary"): fields.Add "RankRate", "IT2"
    CommitSavedFields report, CurrentRevision(report), fields, "Synthetic worn rate"
    group = CreateSummaryGroup("Synthetic reset destination", "EVAL", "E5")
    DoCmd.OpenForm "frmUXMoveGroup", , , , , , CStr(report)
    Forms!frmUXMoveGroup!cboGroup = group
    If Not UXPreviewMoveGroup() Then Err.Raise 5, , "Reset move preview failed"
    first = SaveDirectoryMember(0, 0, "Synthetic Original", "ORIGINAL, SYNTHETIC", "IT2", "")
    second = SaveDirectoryMember(0, 0, "Synthetic Survivor", "SURVIVOR, SYNTHETIC", "IT2", "")
    DoCmd.OpenForm "frmUXDirectoryMerge"
    Forms!frmUXDirectoryMerge!cboSource = first: Forms!frmUXDirectoryMerge!cboTarget = second
    Forms!frmUXDirectoryMerge!txtReason = "Synthetic reviewed duplicate"
    If Not UXDirectoryMergePreview() Then Err.Raise 5, , "Reset merge preview failed"
    PrepareOwnedPreviewReset = "PASS: three owned previews prepared": Exit Function
Failed:
    PrepareOwnedPreviewReset = "FAIL: reset preparation: " & Err.Number & " " & Err.Description
End Function
Public Function VerifyOwnedPreviewReset() As String
    On Error GoTo Failed
    n = 0
    Check Not UXApplyNativeGroups(), "Native import preview survived VBA reset"
    Check Not UXApplyMoveGroup(), "Move preview survived VBA reset"
    Check Not UXDirectoryMergeApply(), "Merge preview survived VBA reset"
    Check DCount("*", "Report") = 1 And DCount("*", "SummaryGroup") = 1, "Reset Apply changed reports/groups"
    Check DCount("*", "GroupMember") = 0 And DCount("*", "DirectoryMerge") = 0 And DCount("*", "ImportReceipt") = 0, "Reset Apply changed links or history"
    Check Not Forms!frmUXNativeGroups!cmdApply.Enabled And Not Forms!frmUXMoveGroup!cmdApply.Enabled And Not Forms!frmUXDirectoryMerge!cmdApply.Enabled, "Reset failure left Apply enabled"
    VerifyOwnedPreviewReset = "PASS: " & n & " real-reset ownership and no-write assertions": Exit Function
Failed:
    VerifyOwnedPreviewReset = "FAIL: reset verification: " & Err.Number & " " & Err.Description
End Function
Public Function TestUXNativeGroups(ByVal source As String, ByVal bad As String, ByVal template As String, ByVal directory As String) As String
    Dim groups As Object, picks As New Collection, key As Variant, item As Object, hash As String, result As Collection, i As Long, id As Long
    Dim members As DAO.Recordset, native As DAO.Database, row As DAO.Recordset, saved As Object, decoded As Object, preview As Object, file As String, owner As Long, second As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC NATIVE GROUP IMPORT", "Importer"
    OpenReviewFile bad, False: Set groups = NAVFITReviewGroupChoices(2000)
    For Each key In groups.Keys: picks.Add CStr(key): hash = groups(key)("SourceHash"): Next
    Check RejectBatch(picks, hash), "Invalid last partition was accepted"
    Check InStr(rejection, "digits only") > 0, "Failure did not reach the last invalid partition: " & rejection
    Check DCount("*", "SummaryGroup") = 0 And DCount("*", "Report") = 0 And DCount("*", "ImportReceipt") = 0, "Failed batch left partial groups/reports/receipts"
    Check batchOutcome = "Rolled back", "Group import failure did not confirm its rollback outcome"
    CloseReviewFile
    Check UXChooseNativeGroupFile(source), "Source chooser adapter failed"
    Check Forms!frmUXNativeGroups!lstGroups.ListCount = 3, "Initial discovery hid groups before century selection: " & Forms!frmUXNativeGroups!lblStatus.Caption
    Check UXLoadNativeGroups(), "Groups did not load: " & Forms!frmUXNativeGroups!lblStatus.Caption
    Check Forms!frmUXNativeGroups!lstGroups.ListCount = 3, "Native group partitions missing"
    Forms!frmUXNativeGroups!lstGroups.Selected(0) = True
    Check UXPreviewNativeGroups(), "Selected group preview failed: " & Forms!frmUXNativeGroups!lblStatus.Caption
    Check InStr(Forms!frmUXNativeGroups!txtPreview, "After:") > 0 And InStr(Forms!frmUXNativeGroups!txtPreview, "NEW GROUP:") > 0, "Full values/group details absent"
    owner = ActiveContributorID(): second = AddContributor("Second synthetic group importer")
    SwitchContributor second
    Check Not UXApplyNativeGroups(), "Another contributor applied the reviewed native groups"
    Check DCount("*", "Report") = 0 And DCount("*", "SummaryGroup") = 0, "Rejected preview ownership wrote data"
    SwitchContributor owner
    Check Not UXApplyNativeGroups(), "Rejected ownership retained an actionable preview"
    Check UXPreviewNativeGroups(), "Original contributor could not preview again"
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
