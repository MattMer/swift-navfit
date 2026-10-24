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
Private exportTemplate As String
Private groupExportToken As String
Private Function NativeExportTemplate() As String
    Dim dom As Object, node As Object, stream As Object
    On Error GoTo Failed
    exportTemplate = Environ$("TEMP") & "\SwiftNAVFIT-template-" & NewIdentity() & ".accdb"
    Set dom = CreateObject("MSXML2.DOMDocument.6.0"): Set node = dom.createElement("template")
    node.DataType = "bin.base64": node.Text = NAVFITTemplateData()
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 1: stream.Open
    stream.Write node.nodeTypedValue: stream.SaveToFile exportTemplate, 1: stream.Close
    If SourceFileHash(exportTemplate) <> NAVFITTemplateHash() Then Err.Raise 5, "NativeExportTemplate", "The internal NAVFIT template could not be verified."
    NativeExportTemplate = exportTemplate: Exit Function
Failed:
    Dim primary As Object
    Set primary = SwiftFailure("modUXNAVFIT", "NativeExportTemplate", "Prepare internal template", Err.Number, Err.Description, Err.Source)
    If Not stream Is Nothing Then SwiftClose stream, primary, "template stream"
    On Error GoTo 0
    SwiftRaise primary
End Function
Public Function UXCloseNAVFITExport() As Boolean
    On Error GoTo Finished
    groupExportToken = ""
    NAVFITClearReturned False
    If Len(exportTemplate) > 0 Then If Len(Dir$(exportTemplate)) > 0 Then Kill exportTemplate
Finished:
    exportTemplate = "": UXCloseNAVFITExport = True
End Function
Public Function UXChooseNAVFITGroupFolder() As Boolean
    With Application.FileDialog(4)
        .Title = "Choose the NAVFIT98A export directory"
        If .Show <> -1 Then Exit Function
        Forms!frmNAVFITExchange!txtExportPath = .SelectedItems(1)
    End With
    UXChooseNAVFITGroupFolder = UXPreviewNAVFITGroup()
End Function
Public Function UXExportSelectedGroup() As Boolean
    Dim id As Long
    id = Nz(Forms!frmUXGroups!lstGroups, 0)
    If id = 0 Then Forms!frmUXGroups!lblInfo.Caption = "Select a summary group to export.": Exit Function
    UXExportSelectedGroup = UXOpenNAVFITGroupExport(id)
End Function





Public Function UXOpenNAVFITGroupExport(Optional ByVal groupID As Long = 0) As Boolean
    Dim id As Long
    id = groupID: If id = 0 Then id = Nz(Forms!frmMajorWorkbench!cboGroup, 0)
    If id = 0 Then Exit Function
    Call UXOpenNAVFITExchange
    Forms!frmNAVFITExchange!cboGroup = id
    Forms!frmNAVFITExchange!lblGroup.Caption = "Summary group: " & Nz(DLookup("Title", "SummaryGroup", "GroupID=" & id), "")
    Forms!frmNAVFITExchange!chkPSG = True
    UXOpenNAVFITGroupExport = UXPreviewNAVFITGroup()
End Function
Public Function UXPreviewNAVFITGroup() As Boolean
    Dim preview As Object
    On Error GoTo Failed
    groupExportToken = "": Forms!frmNAVFITExchange!cmdExportGroup.Enabled = False
    Set preview = NAVFITGroupPreview(Nz(Forms!frmNAVFITExchange!cboGroup, 0), Nz(Forms!frmNAVFITExchange!chkPSG, False))
    groupExportToken = preview("Token")
    Forms!frmNAVFITExchange!txtStatus = DCount("*", "GroupMember", "GroupID=" & Nz(Forms!frmNAVFITExchange!cboGroup, 0)) & " saved reports will be exported. Findings do not block export. Choose the export directory, then Export."
    Forms!frmNAVFITExchange!cmdExportGroup.Enabled = True: UXPreviewNAVFITGroup = True: Exit Function
Failed:
    Forms!frmNAVFITExchange!txtStatus = "Group preview failed: " & Err.Description
End Function
Public Function UXExportNAVFITGroup() As Boolean
    Dim findings As Collection
    On Error GoTo Failed
    If Not UXPreviewNAVFITGroup() Then Exit Function
    With Forms!frmNAVFITExchange
        !txtExportPath = UXExportDestination(Nz(!txtExportPath, ""), Nz(!cboGroup, 0), 0)
        Call NAVFITExportGroup(Nz(!cboGroup, 0), Nz(!txtExportPath, ""), Nz(!txtTemplate, ""), groupExportToken, Nz(!chkPSG, False))
        Set findings = NAVFITLastGroupReadiness()
        !txtStatus = "Export complete: " & !txtExportPath & vbCrLf & "Open the .accdb file in NAVFIT98A. All saved group members are included. Pending edits remain in Swift."
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
    text = "NEW " & values("FormKind") & " REPORT" & vbCrLf & "" & NAVFITMapRows(values("FormKind")).Count & " saved official fields are mapped. All 128 native report columns, folder defaults and summary values are retained in provenance. Native-only values are preserved for export; they are not editable in Swift's current editor. Native validation is not a local approval." & vbCrLf
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
    If nativePreview.OwnerProfileID <> ActiveContributorID() Then Err.Raise vbObjectError + 9203, "modUXNAVFIT.UXApplyFullNAVFIT", "Preview this import as the current contributor before Apply."
    If nativePreview.Member <> Nz(Forms!frmImportReview!txtImportMember, "") Then Err.Raise vbObjectError + 5500, "modUXNAVFIT.UXApplyFullNAVFIT", "Display name changed. Preview again."
    If nativePreview.SourceID = 0 Or nativePreview.SourceID <> Nz(Forms!frmImportReview!lstImportReports, 0) Then Err.Raise vbObjectError + 5500, "modUXNAVFIT.UXApplyFullNAVFIT", "Preview the current source before Apply."
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
    'The exchange dialog exports saved values and must not save/discard an editor.
    If CurrentProject.AllForms("frmNAVFITExchange").IsLoaded Then DoCmd.Close acForm, "frmNAVFITExchange", acSaveNo
    DoCmd.OpenForm "frmNAVFITExchange"
    Forms!frmNAVFITExchange!txtTemplate = NativeExportTemplate()
    Forms!frmNAVFITExchange!cboReport.RowSource = "SELECT ReportID, MemberName & ' | ' & FormKind & ' | ' & Nz(Format(PeriodEnd,'yyyy-mm-dd'),'No end date') & IIf(IsArchived,' | Archived','') FROM Report ORDER BY MemberName"
    Forms!frmNAVFITExchange!cboGroup.RowSource = "SELECT GroupID,Title & ' | ' & FormKind & ' | ' & Paygrade & IIf(IsArchived,' | Archived','') FROM SummaryGroup ORDER BY Title"
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
        !txtStatus = "Exporting the saved revision. Findings do not block export; pending edits are not included."
        !txtExportPath = UXExportDestination(Nz(!txtExportPath, ""), 0, Nz(!cboReport, 0))
        Call NAVFITExportReport(Nz(!cboReport, 0), Nz(!txtExportPath, ""), Nz(!txtTemplate, ""))
        Set findings = NAVFITLastReportReadiness()
        !txtStatus = "Export complete: " & !txtExportPath & vbCrLf & "Findings do not block export. Export is not approval, signature or submission. Pending edits are not included. Open All report readiness issues to inspect findings. Keep the .swift-baseline.accdb file in Swift. Send only the NAVFIT .accdb file. This export contains one report; it does not certify a complete summary group." & vbCrLf & vbCrLf & ReadinessSummary(findings)
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
    Dim id As Long, outcome As String, message As String
    outcome = "Not started"
    On Error GoTo Failed
    id = NAVFITApplyReturned(outcome)
    Forms!frmNAVFITExchange!cmdApply.Enabled = False
    Forms!frmNAVFITExchange!txtStatus = "Return applied to " & UXMemberLabel(id) & ". Provenance and normal revision history retained."
    UXApplyNAVFITReturn = True: Exit Function
Failed:
    message = Err.Description
    Select Case outcome
    Case "Committed"
        Forms!frmNAVFITExchange!cmdApply.Enabled = False
        NAVFITClearReturned False
        Forms!frmNAVFITExchange!txtStatus = "Return data was committed; finalization failed. Do not re-apply. Close and reopen this workspace. " & message
    Case "Not started", "Rolled back"
        If outcome = "Rolled back" Then Forms!frmNAVFITExchange!cmdApply.Enabled = False
        Forms!frmNAVFITExchange!txtStatus = "Return " & LCase$(outcome) & ": " & message & " Preview again before retrying."
    Case Else
        Forms!frmNAVFITExchange!cmdApply.Enabled = False
        NAVFITClearReturned False
        Forms!frmNAVFITExchange!txtStatus = "Return outcome could not be confirmed. Close and reopen; inspect saved data before retrying. " & message
    End Select
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
    On Error GoTo Failed
    UXCancelNAVFITReturn = NAVFITClearReturned()
    On Error GoTo 0
    Forms!frmNAVFITExchange!lstChoices.RowSource = ""
    Forms!frmNAVFITExchange!txtDetail = "": Forms!frmNAVFITExchange!cmdApply.Enabled = False
    Forms!frmNAVFITExchange!txtStatus = "Saved return preview cancelled. No report values changed."
    Exit Function
Failed:
    Forms!frmNAVFITExchange!cmdApply.Enabled = False
    Forms!frmNAVFITExchange!txtStatus = "Cancel unavailable: " & Err.Description & " Saved choices are retained. Resume before cancelling."
End Function
