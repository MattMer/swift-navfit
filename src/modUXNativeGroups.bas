Attribute VB_Name = "modUXNativeGroups"
Option Explicit
Private sourcePath As String
Private choices As Object
Private selectedPreview As Collection
Private previewHash As String
Private previewCentury As Long
Private gradeOverrides As Object
Private Function SelectedGroups() As Collection
    Dim result As New Collection, i As Long, f As Access.Form
    Set f = Forms!frmUXNativeGroups
    For i = 0 To f!lstGroups.ListCount - 1
        If f!lstGroups.Selected(i) Then result.Add CStr(f!lstGroups.Column(0, i))
    Next
    Set SelectedGroups = result
End Function
Private Function SelectionKey(ByVal items As Collection) As String
    Dim key As Variant
    For Each key In items: SelectionKey = SelectionKey & CStr(key) & ";": Next
End Function
Public Function UXChooseNativeGroups() As Boolean
    On Error GoTo Failed
    With Application.FileDialog(3)
        .Title = "Choose NAVFIT database for group import": .AllowMultiSelect = False
        .Filters.Clear: .Filters.Add "NAVFIT database", "*.accdb"
        If .Show <> -1 Then Exit Function
        UXChooseNativeGroupFile .SelectedItems(1)
    End With
    UXChooseNativeGroups = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "NAVFIT group import"
End Function
Public Function UXChooseNativeGroupFile(ByVal path As String) As Boolean
    OpenReviewFile path, False: sourcePath = path
    Set gradeOverrides = CreateObject("Scripting.Dictionary")
    DoCmd.OpenForm "frmUXNativeGroups"
    Call UXLoadNativeGroups
    UXChooseNativeGroupFile = True
End Function
Public Function UXLoadNativeGroups() As Boolean
    Dim f As Access.Form, key As Variant, item As Object
    On Error GoTo Failed
    Set f = Forms!frmUXNativeGroups: Set selectedPreview = Nothing: previewHash = "": f!cmdApply.Enabled = False
    f!txtSource = sourcePath
    Set choices = NAVFITReviewGroupChoices(2000, gradeOverrides, True)
    f!lstGroups.RowSourceType = "Value List": f!lstGroups.RowSource = ""
    For Each key In choices.Keys
        Set item = choices(key)
        f!lstGroups.AddItem key & ";" & Replace(item("Title"), ";", ",") & ";" & item("Kind") & ";" & item("Paygrade") & ";" & item("Count")
    Next
    f!lblStatus.Caption = ReviewUpgradeMessage() & "Choose groups, then Preview. A ? paygrade means the source rank/rate is missing: select those groups and assign a compatible paygrade below. Empty folders contain no reports to import."
    If choices.Count = 0 Then f!lblStatus.Caption = "This database contains no populated report groups. Empty folders have no reports to import."
    UXLoadNativeGroups = True: Exit Function
Failed:
    Forms!frmUXNativeGroups!lblStatus.Caption = Err.Description
End Function
Public Function UXAssignNativeGroupGrade() As Boolean
    Dim picks As Collection, key As Variant, item As Object, grade As String, parts As Variant, targets As Object, i As Long
    On Error GoTo Failed
    Set picks = SelectedGroups(): If picks.Count = 0 Then Err.Raise 5, , "Select a group with unresolved source rank/rate first."
    grade = UCase$(Trim$(Nz(Forms!frmUXNativeGroups!cboImportGrade, "")))
    'Validate the entire selection before changing any local grouping metadata.
    For Each key In picks
        Set item = choices(CStr(key))
        If item("UnresolvedCount") = 0 Then Err.Raise 5, , "Select only groups containing unresolved source rank/rate."
        If Not NativeImportGradeMatches(item("Kind"), grade) Then Err.Raise 5, , "Choose a compatible paygrade for " & item("Kind") & "."
    Next
    Set targets = CreateObject("Scripting.Dictionary")
    For Each key In picks
        parts = Split(CStr(key), "|")
        gradeOverrides(parts(0) & "|" & parts(1) & "|?") = grade
        targets(parts(0) & "|" & parts(1) & "|" & grade) = True
    Next
    UXLoadNativeGroups
    For i = 0 To Forms!frmUXNativeGroups!lstGroups.ListCount - 1
        Forms!frmUXNativeGroups!lstGroups.Selected(i) = targets.Exists(CStr(Forms!frmUXNativeGroups!lstGroups.Column(0, i)))
    Next
    Forms!frmUXNativeGroups!lblStatus.Caption = "Local group paygrade assigned. Source files and report rank/rate values are unchanged. Preview the complete selected roster before applying."
    UXAssignNativeGroupGrade = True: Exit Function
Failed:
    Forms!frmUXNativeGroups!lblStatus.Caption = Err.Description
End Function
Public Function UXNativeGroupSelectionChanged() As Boolean
    Set selectedPreview = Nothing: previewHash = "": Forms!frmUXNativeGroups!cmdApply.Enabled = False
    UXNativeGroupSelectionChanged = True
End Function
Public Function UXPreviewNativeGroups() As Boolean
    Dim key As Variant, item As Object, nativeID As Variant, values As Object, field As Variant, text As String, picks As Collection
    On Error GoTo Failed
    UXNativeGroupSelectionChanged
    Set picks = SelectedGroups(): If picks.Count = 0 Then Err.Raise 5, , "Select at least one source group."
    previewCentury = 2000
    Set choices = NAVFITReviewGroupChoices(previewCentury, gradeOverrides)
    For Each key In picks
        Set item = choices(CStr(key)): previewHash = item("SourceHash")
        If item("Blocked") <> "" Then Err.Raise 5, , item("Title") & ": " & item("Blocked")
        text = text & vbCrLf & "NEW GROUP: " & item("Title") & " - " & item("Kind") & " " & item("Paygrade") & " (" & item("Count") & " reports)" & vbCrLf
        If item("UnresolvedCount") > 0 Then text = text & "Group paygrade assigned locally for " & item("UnresolvedCount") & " unresolved source rates. Report rank/rate values remain unchanged and need completion after import." & vbCrLf
        For Each field In item("Defaults").Keys
            text = text & "Default " & UXFieldCaption(CStr(field), item("Kind")) & ": " & NAVFITDisplay(item("Defaults")(field)) & vbCrLf
        Next
        For Each nativeID In Split(item("Reports"), ",")
            Set values = FullNAVFITImportValues(CLng(nativeID), previewCentury)
            If DCount("*", "ImportReceipt", "SourceHash=" & QuoteSQL(previewHash) & " AND ExternalReportID=" & CLng(nativeID)) > 0 Then Err.Raise 5, , "Selected member already imported. Use returned-file reconciliation instead."
            text = text & vbCrLf & "NEW REPORT: " & NAVFITDisplay(values("OfficialName")) & vbCrLf
            For Each field In values.Keys
                If field <> "SourceHash" Then text = text & UXFieldCaption(CStr(field), item("Kind")) & vbCrLf & "Before: [new report]" & vbCrLf & "After: " & NAVFITDisplay(values(field)) & vbCrLf
            Next
        Next
    Next
    Set selectedPreview = picks
    Forms!frmUXNativeGroups!txtPreview = text & vbCrLf & "Native dates use their calendar date; original timestamps are retained in provenance. Native-only values and summary/folder columns are retained in report provenance. Import does not establish roster completeness or Navy compliance. No existing draft is overwritten."
    Forms!frmUXNativeGroups!lblStatus.Caption = "Preview ready. Apply creates selected groups and reports together."
    Forms!frmUXNativeGroups!cmdApply.Enabled = True
    UXPreviewNativeGroups = True: Exit Function
Failed:
    Forms!frmUXNativeGroups!lblStatus.Caption = Err.Description: Forms!frmUXNativeGroups!cmdApply.Enabled = False
End Function
Public Function UXApplyNativeGroups() As Boolean
    Dim result As Collection, picks As Collection
    On Error GoTo Failed
    If selectedPreview Is Nothing Then Err.Raise 5, , "Preview the selected groups first."
    Set picks = SelectedGroups()
    If SelectionKey(picks) <> SelectionKey(selectedPreview) Then Err.Raise 5, , "Selection changed. Preview again."
    Set result = ImportNAVFITGroups(picks, previewHash, previewCentury, gradeOverrides)
    Forms!frmUXNativeGroups!cmdPreview.SetFocus: Forms!frmUXNativeGroups!cmdApply.Enabled = False
    Set selectedPreview = Nothing
    Forms!frmUXNativeGroups!lblStatus.Caption = "Results: " & result.Count & " groups imported with all their selected reports. Open Summary groups to inspect members and validation."
    UXApplyNativeGroups = True: Exit Function
Failed:
    Forms!frmUXNativeGroups!lblStatus.Caption = Err.Description & " No partial selected-group batch was committed."
    Forms!frmUXNativeGroups!cmdApply.Enabled = False
End Function
Public Function UXCloseNativeGroups() As Boolean
    CloseReviewFile: DoCmd.Close acForm, "frmUXNativeGroups", acSaveNo
    UXCloseNativeGroups = OpenScreen("frmUXImportChoose")
End Function

Public Function UXSelectAllNativeGroups(ByVal selectAll As Boolean) As Boolean
 Dim f As Access.Form, i As Long, firstRow As Long, selectedCount As Long
 On Error GoTo Failed
 Set f = Forms!frmUXNativeGroups
 UXNativeGroupSelectionChanged
 firstRow = IIf(f!lstGroups.ColumnHeads, 1, 0)
 For i = firstRow To f!lstGroups.ListCount - 1
  f!lstGroups.Selected(i) = selectAll
  If selectAll Then selectedCount = selectedCount + 1
 Next
 f!txtPreview = Null
 f!lblStatus.Caption = selectedCount & " groups selected. Preview the selection before importing. Unresolved paygrades still need assignment."
 UXSelectAllNativeGroups = True
 Exit Function
Failed:
 Forms!frmUXNativeGroups!lblStatus.Caption = Err.Description
End Function
