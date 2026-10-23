Attribute VB_Name = "modUXNAVFIT"
Option Compare Binary
Option Explicit
Private Type NativeImportPreviewContext
    SourceID As Long
    Hash As String
    Century As Long
    Created As Long
    Member As String
    OwnerProfileID As Long
End Type
Private nativePreview As NativeImportPreviewContext





Private groupExportToken As String
Public Function UXOpenNAVFITGroupExport(Optional ByVal groupID As Long = 0) As Boolean
    Dim id As Long
    If Not UXLeaveEditor() Then Exit Function
    id = groupID: If id = 0 Then id = Nz(Forms!frmMajorWorkbench!cboGroup, 0)
    If id = 0 Then Exit Function
    Call UXOpenNAVFITExchange
    Forms!frmNAVFITExchange!cboGroup = id
    UXOpenNAVFITGroupExport = True
End Function
Public Function UXPreviewNAVFITGroup() As Boolean
    Dim preview As Object
    On Error GoTo Failed
    groupExportToken = "": Forms!frmNAVFITExchange!cmdExportGroup.Enabled = False
    Set preview = NAVFITGroupPreview(Nz(Forms!frmNAVFITExchange!cboGroup, 0), Nz(Forms!frmNAVFITExchange!chkPSG, False))
    groupExportToken = preview("Token"): Forms!frmNAVFITExchange!txtStatus = preview("Text")
    Forms!frmNAVFITExchange!cmdExportGroup.Enabled = True: UXPreviewNAVFITGroup = True: Exit Function
Failed:
    Forms!frmNAVFITExchange!txtStatus = "Group preview failed: " & Err.Description
End Function
Public Function UXExportNAVFITGroup() As Boolean
    Dim findings As Collection
    On Error GoTo Failed
    With Forms!frmNAVFITExchange
        !txtExportPath = UXExportDestination(Nz(!txtExportPath, ""), Nz(!cboGroup, 0), 0)
        Call NAVFITExportGroup(Nz(!cboGroup, 0), Nz(!txtExportPath, ""), Nz(!txtTemplate, ""), groupExportToken, Nz(!chkPSG, False))
        Set findings = NAVFITLastGroupReadiness()
        !txtStatus = "Group export complete: " & !txtExportPath & vbCrLf & "Open this file in NAVFIT98A, select the group folder, validate its reports and use NAVFIT's Print Folder/Print Summary actions. Keep the group baseline locally. Group-return reconciliation is pending."
        !txtStatus = !txtStatus & vbCrLf & vbCrLf & ReadinessSummary(findings)
        !cmdExportGroup.Enabled = False
    End With
    groupExportToken = "": UXExportNAVFITGroup = True: Exit Function
Failed:
    Forms!frmNAVFITExchange!txtStatus = "Group export failed: " & Err.Description
End Function
Public Function UXPreviewFullNAVFIT() As Boolean
    Dim values As Object, key As Variant, text As String, century As Long, id As Long
    On Error GoTo Failed
    nativePreview.SourceID = 0: nativePreview.Hash = "": nativePreview.Created = 0
    century = 2000: id = Nz(Forms!frmImportReview!lstImportReports, 0)
    nativePreview.Member = Nz(Forms!frmImportReview!txtImportMember, "")
    Set values = FullNAVFITImportValues(id, century)
    text = "NEW " & values("FormKind") & " REPORT" & vbCrLf & "74 saved official fields are mapped. All 128 native report columns, folder defaults and summary values are retained in provenance. Native-only values are preserved for export; they are not editable in Swift's current editor. Native validation is not a local approval." & vbCrLf
    text = text & vbCrLf & "Native dates use their calendar date; original timestamps are retained in provenance." & vbCrLf
    text = text & vbCrLf & "Local display name: " & IIf(nativePreview.Member = "", NAVFITDisplay(values("OfficialName")), nativePreview.Member) & vbCrLf
    For Each key In values.Keys
        If key <> "SourceHash" And key <> "FormKind" Then text = text & vbCrLf & UXFieldCaption(CStr(key), values("FormKind")) & vbCrLf & "Before: [new report]" & vbCrLf & "After: " & NAVFITDisplay(values(key)) & vbCrLf
    Next
    With Forms!frmImportReview
        !cmdPreview.SetFocus: !uxPreview.Visible = True: !uxPreview = text
        !lstImportFields.Visible = False: !txtFieldValue.Visible = False: !lblFieldColumns.Visible = False
        !cmdImport.Enabled = True: !uxState.Caption = "3. Apply - review all mapped values before creating a report."
    End With
    nativePreview.SourceID = id: nativePreview.Hash = values("SourceHash"): nativePreview.Century = century
    nativePreview.OwnerProfileID = ActiveContributorID()
    UXPreviewFullNAVFIT = True: Exit Function
Failed:
    Forms!frmImportReview!cmdImport.Enabled = False: Forms!frmImportReview!uxState.Caption = "2. Preview - " & Err.Description
End Function
Public Function UXApplyFullNAVFIT() As Boolean
    On Error GoTo Failed
    If nativePreview.OwnerProfileID <> ActiveContributorID() Then Err.Raise vbObjectError + 9203, , "Preview this import as the current contributor before Apply."
    If nativePreview.Member <> Nz(Forms!frmImportReview!txtImportMember, "") Then Err.Raise vbObjectError + 5500, , "Display name changed. Preview again."
    If nativePreview.SourceID = 0 Or nativePreview.SourceID <> Nz(Forms!frmImportReview!lstImportReports, 0) Then Err.Raise vbObjectError + 5500, , "Preview the current source before Apply."
    nativePreview.Created = ImportReviewedReport(nativePreview.SourceID, nativePreview.Hash, nativePreview.Century, nativePreview.Member): nativePreview.SourceID = 0: nativePreview.Hash = ""
    With Forms!frmImportReview
        !cmdPreview.SetFocus: !cmdImport.Enabled = False: !cmdOpenResult.Enabled = True
        !uxState.Caption = "4. Results - created " & UXMemberLabel(nativePreview.Created) & ". Native source values retained."
    End With
    UXApplyFullNAVFIT = True: Exit Function
Failed:
    nativePreview.SourceID = 0: nativePreview.Hash = "": Forms!frmImportReview!cmdImport.Enabled = False
    Forms!frmImportReview!uxState.Caption = Err.Description & " Preview again before retrying."
End Function
Public Function UXOpenFullNAVFITResult() As Boolean
    If nativePreview.Created = 0 Then Exit Function
    If Not UXLeaveEditor() Then Exit Function
    DoCmd.OpenForm "frmEditor", , , , , , CStr(nativePreview.Created)
    UXOpenFullNAVFITResult = True
End Function
Public Function UXOpenNAVFITExchange() As Boolean
    If Not UXLeaveEditor() Then Exit Function
    DoCmd.OpenForm "frmNAVFITExchange"
    Forms!frmNAVFITExchange!cboReport.RowSource = "SELECT ReportID, MemberName & ' | ' & FormKind & ' | ' & Nz(Format(PeriodEnd,'yyyy-mm-dd'),'No end date') FROM Report WHERE IsArchived=False ORDER BY MemberName"
    Forms!frmNAVFITExchange!cboGroup.RowSource = "SELECT GroupID,Title & ' | ' & FormKind & ' | ' & Paygrade FROM SummaryGroup WHERE IsArchived=False ORDER BY Title"
    UXOpenNAVFITExchange = True
End Function
Public Function UXChooseNAVFITPath(ByVal controlName As String, ByVal savePath As Boolean) As Boolean
    With Application.FileDialog(IIf(savePath, 4, 3))
        .Title = "Choose " & controlName
        If .Show <> -1 Then Exit Function
        Forms!frmNAVFITExchange.Controls(controlName) = .SelectedItems(1)
    End With
    Call UXNAVFITInputChanged
    UXChooseNAVFITPath = True
End Function
Public Function UXExportNAVFIT() As Boolean
    Dim findings As Collection
    On Error GoTo Failed
    With Forms!frmNAVFITExchange
        !txtStatus = "Checking saved report readiness and preparing export. Open All report readiness issues to inspect findings before exporting."
        !txtExportPath = UXExportDestination(Nz(!txtExportPath, ""), 0, Nz(!cboReport, 0))
        Call NAVFITExportReport(Nz(!cboReport, 0), Nz(!txtExportPath, ""), Nz(!txtTemplate, ""))
        Set findings = NAVFITLastReportReadiness()
        !txtStatus = "Export complete: " & !txtExportPath & vbCrLf & "Keep the .swift-baseline.accdb file in Swift. Send only the NAVFIT .accdb file. This export contains one report; it does not certify a complete summary group." & vbCrLf & vbCrLf & ReadinessSummary(findings)
        !txtBaseline = !txtExportPath & ".swift-baseline.accdb"
    End With
    UXExportNAVFIT = True: Exit Function
Failed:
    Forms!frmNAVFITExchange!txtStatus = "Export failed: " & Err.Description
End Function
Public Function UXPreviewNAVFITReturn() As Boolean
    Dim row As Variant, parts As Variant, returnChoices As Collection
    On Error GoTo Failed
    With Forms!frmNAVFITExchange
        !cmdApply.Enabled = False
        Set returnChoices = NAVFITPreviewReturned(Nz(!txtReturned, ""), Nz(!txtBaseline, ""), 2000)
        !lstChoices.RowSourceType = "Value List": !lstChoices.RowSource = ""
        For Each row In returnChoices
            parts = Split(row, "|")
            !lstChoices.AddItem parts(0) & ";" & Replace(NAVFITReturnedCaption(CStr(parts(0))), ";", ",") & ";" & parts(1)
        Next
        !txtStatus = "Preview ready. Select each field to inspect the complete baseline, current and returned values. Resolve every Conflict. Apply also accepts the displayed native-only source changes; cancel this return if those are unwanted."
        !cmdApply.Enabled = True
    End With
    UXPreviewNAVFITReturn = True: Exit Function
Failed:
    Forms!frmNAVFITExchange!txtStatus = "Preview failed: " & Err.Description
End Function
Public Function UXNAVFITReturnDetail() As Boolean
    If IsNull(Forms!frmNAVFITExchange!lstChoices) Then Exit Function
    Forms!frmNAVFITExchange!cmdKeep.Enabled = (Left$(CStr(Forms!frmNAVFITExchange!lstChoices), 7) <> "Native:")
    Forms!frmNAVFITExchange!cmdAccept.Enabled = Forms!frmNAVFITExchange!cmdKeep.Enabled
    Forms!frmNAVFITExchange!txtDetail = NAVFITReturnedDetails(CStr(Forms!frmNAVFITExchange!lstChoices))
    UXNAVFITReturnDetail = True
End Function
Public Function UXResolveNAVFITReturn(ByVal choice As String) As Boolean
    Dim selected As String, row As Variant, parts As Variant, returnChoices As Collection
    On Error GoTo Failed
    selected = Nz(Forms!frmNAVFITExchange!lstChoices, "")
    NAVFITResolveReturned selected, choice
    Set returnChoices = NAVFITReturnedRows()
    With Forms!frmNAVFITExchange!lstChoices
        .RowSource = ""
        For Each row In returnChoices
            parts = Split(row, "|")
            .AddItem parts(0) & ";" & Replace(NAVFITReturnedCaption(CStr(parts(0))), ";", ",") & ";" & NAVFITReturnedDecision(CStr(parts(0)))
        Next
        .Value = selected
    End With
    UXResolveNAVFITReturn = UXNAVFITReturnDetail(): Exit Function
Failed:
    Forms!frmNAVFITExchange!txtStatus = Err.Description
End Function
Public Function UXApplyNAVFITReturn() As Boolean
    Dim id As Long
    On Error GoTo Failed
    id = NAVFITApplyReturned()
    Forms!frmNAVFITExchange!cmdApply.Enabled = False
    Forms!frmNAVFITExchange!txtStatus = "Return applied to " & UXMemberLabel(id) & ". Provenance and normal revision history retained."
    UXApplyNAVFITReturn = True: Exit Function
Failed:
    Forms!frmNAVFITExchange!txtStatus = "Nothing applied: " & Err.Description
End Function


Public Function UXNAVFITInputChanged() As Boolean
    groupExportToken = "": Forms!frmNAVFITExchange!cmdExportGroup.Enabled = False
    NAVFITClearReturned False
    Forms!frmNAVFITExchange!cmdApply.Enabled = False
    Forms!frmNAVFITExchange!lstChoices.RowSource = ""
    Forms!frmNAVFITExchange!txtDetail = ""
    Forms!frmNAVFITExchange!txtStatus = "Choices changed. Preview again before applying a returned file."
    UXNAVFITInputChanged = True
End Function
Public Function UXResumeNAVFITReturn() As Boolean
    Dim rows As Collection, row As Variant, parts As Variant, saved As Object
    On Error GoTo Failed
    Forms!frmNAVFITExchange!cmdApply.Enabled = False
    Set saved = WizardPreviewLoad("NAVFITReturn")
    Set rows = NAVFITResumeReturned()
    With Forms!frmNAVFITExchange
        !txtReturned = saved("SourcePath"): !txtBaseline = saved("BaselinePath")
        !lstChoices.RowSourceType = "Value List"
        !lstChoices.RowSource = ""
        For Each row In rows
            parts = Split(row, "|")
            !lstChoices.AddItem parts(0) & ";" & Replace(NAVFITReturnedCaption(CStr(parts(0))), ";", ",") & ";" & parts(1)
        Next
        !txtStatus = "Saved choices resumed after source, baseline, contributor and target checks. Review before Apply. Nothing applied automatically."
        !cmdApply.Enabled = True
    End With
    UXResumeNAVFITReturn = True: Exit Function
Failed:
    Forms!frmNAVFITExchange!txtStatus = "Resume unavailable: " & Err.Description
End Function
Public Function UXCancelNAVFITReturn() As Boolean
    UXCancelNAVFITReturn = NAVFITClearReturned()
    Forms!frmNAVFITExchange!lstChoices.RowSource = ""
    Forms!frmNAVFITExchange!txtDetail = "": Forms!frmNAVFITExchange!cmdApply.Enabled = False
    Forms!frmNAVFITExchange!txtStatus = "Saved return preview cancelled. No report values changed."
End Function
