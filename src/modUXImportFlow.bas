Attribute VB_Name = "modUXImportFlow"
Option Compare Binary
Option Explicit
Private pdfSignature As String
Private pdfCreated As Long
Private narrativeSource As Long
Private narrativeHash As String
Private narrativeMember As String
Private narrativeCreated As Long
Private Function ImportValue(ByVal value As Variant) As String
    If IsNull(value) Or Len(Nz(value, "")) = 0 Then
        ImportValue = "[blank]"
    ElseIf VarType(value) = vbDate Then
        ImportValue = Format$(value, "yyyy-mm-dd")
    Else
        ImportValue = CStr(value)
    End If
End Function
Private Sub ImportState(ByVal formName As String, ByVal step As String, ByVal message As String, ByVal ready As Boolean)
    Dim f As Access.Form
    Set f = Forms(formName)
    If Not ready Then f!cmdPreview.SetFocus
    f!uxState.Caption = step & " - " & message
    f!cmdImport.Enabled = ready
End Sub
Public Function UXOpenImportChooser() As Boolean
    If Not UXLeaveEditor() Then Exit Function
    DoCmd.OpenForm "frmUXImportChoose"
    UXOpenImportChooser = True
End Function
Public Function UXPDFStart() As Boolean
    Call PDFClearPreview
    pdfSignature = "": pdfCreated = 0
    Forms!frmPDFImport!cmdImport.Enabled = False: Forms!frmPDFImport!cmdOpenResult.Enabled = False
    Forms!frmPDFImport!uxState.Caption = "1. Choose - select a supported structured PDF."
    Forms!frmPDFImport!cboTarget.RowSource = "SELECT ReportID, MemberName & ' | ' & FormKind & ' | ' & Nz(Format(PeriodEnd,'yyyy-mm-dd'),'No end date') AS Description FROM Report WHERE IsArchived=False AND WorkflowState='Draft' ORDER BY MemberName"
    Forms!frmPDFImport!lblStatus.Caption = PDFExtractionStatus()
    UXPDFStart = True
End Function
Public Function UXChoosePDFFile(ByVal path As String) As Boolean
    Dim result As Object
    On Error GoTo Failed
    pdfSignature = "": pdfCreated = 0
    Forms!frmPDFImport!cmdOpenResult.Enabled = False
    Set result = PreviewStructuredPDF(path)
    Forms!frmPDFImport!txtPreview = "Source: " & path & vbCrLf & "Report type: " & result("FormKind") & vbCrLf & "Choose the identifier interpretation, then Preview mapping. Two-digit years use 20xx."
    Call ImportState("frmPDFImport", "2. Preview", "Review mapping before applying. No draft created.", False)
    UXChoosePDFFile = True: Exit Function
Failed:
    Call ImportState("frmPDFImport", "1. Choose", Err.Description, False)
End Function
Public Function UXPDFMappingChanged() As Boolean
    pdfSignature = "": pdfCreated = 0
    Call ImportState("frmPDFImport", "2. Preview", "Mapping choices changed. Preview again before Apply.", False)
    Forms!frmPDFImport!cmdOpenResult.Enabled = False
    UXPDFMappingChanged = True
End Function
Public Function UXPDFLegacyChanged() As Boolean
    With Forms!frmPDFImport
        If !chkLegacy Then
            !cboIdentifier.RowSource = "SSN;Retain only;DoDID"
        Else
            If Nz(!cboIdentifier, "") = "DoDID" Then !cboIdentifier = "SSN"
            !cboIdentifier.RowSource = "SSN;Retain only"
        End If
    End With
    UXPDFLegacyChanged = UXPDFMappingChanged()
End Function
Public Function UXPreviewPDFImport() As Boolean
    Dim changes As Object, incoming As Object, before As Object, key As Variant, text As String, kind As String, target As Long, mode As String, outcome As String
    On Error GoTo Failed
    pdfSignature = "": pdfCreated = 0
    Forms!frmPDFImport!cmdPreview.SetFocus
    Forms!frmPDFImport!cmdOpenResult.Enabled = False
    If Nz(Forms!frmPDFImport!cboIdentifier, "") = "DoDID" And Not Nz(Forms!frmPDFImport!chkLegacy, False) Then Err.Raise vbObjectError + 4903, , "Enable legacy compatibility explicitly before mapping DoDID."
    Set incoming = PDFMappedValues("20", Nz(Forms!frmPDFImport!cboIdentifier, ""))
    kind = PDFImportKind()
    target = Nz(Forms!frmPDFImport!cboTarget, 0): mode = Nz(Forms!frmPDFImport!cboMergeMode, "Fill empty fields")
    If target > 0 Then
        Set changes = PDFPrepareMerge("20", Nz(Forms!frmPDFImport!cboIdentifier, ""), target, mode)
        Set before = SavedFields(target)
        text = "MERGE INTO " & UXMemberLabel(target) & " - " & mode & vbCrLf
    Else
        Set changes = incoming
        text = "NEW " & kind & " DRAFT" & vbCrLf
    End If
    text = text & "Supported catalogue: " & PDFMapRows(kind).Count & " saved fields. Only present nonblank values apply. Missing/blank PDF fields preserve existing values. Unmapped content stays in retained source XML. Signatures, calculated averages, command achievements, combined addresses and unnamed checkboxes are not imported." & vbCrLf
    For Each key In incoming.Keys
        If target = 0 Then
            text = text & vbCrLf & UXFieldCaption(CStr(key), kind) & vbCrLf & "Before: [new draft / no prior value]" & vbCrLf & "After: " & ImportValue(incoming(key)) & vbCrLf
        Else
            outcome = "Keep existing: " & ImportValue(before(key))
            If changes.Exists(key) Then outcome = "After: " & ImportValue(changes(key))
            text = text & vbCrLf & UXFieldCaption(CStr(key), kind) & vbCrLf & "Before: " & ImportValue(before(key)) & vbCrLf & "PDF: " & ImportValue(incoming(key)) & vbCrLf & outcome & vbCrLf
        End If
    Next
    If changes.Count = 0 Then Err.Raise vbObjectError + 4904, , "No changed values to apply. Choose another merge mode or target."
    Forms!frmPDFImport!txtPreview = text
    pdfSignature = UXPDFSignature()
    Call ImportState("frmPDFImport", "3. Apply", CStr(changes.Count) & " values ready. Review full before/after values; Apply retains source provenance.", True)
    UXPreviewPDFImport = True: Exit Function
Failed:
    Call ImportState("frmPDFImport", "2. Preview", Err.Description, False)
End Function
Public Function UXApplyPDFImport() As Boolean
    On Error GoTo Failed
    If Len(pdfSignature) = 0 Or pdfSignature <> UXPDFSignature() Then Err.Raise vbObjectError + 4901, , "Preview the current mapping before Apply."
    pdfCreated = ImportPDFPreview("20", Nz(Forms!frmPDFImport!cboIdentifier, ""), Nz(Forms!frmPDFImport!cboTarget, 0), Nz(Forms!frmPDFImport!cboMergeMode, "Fill empty fields"))
    pdfSignature = ""
    Call ImportState("frmPDFImport", "4. Results", "Imported into " & UXMemberLabel(pdfCreated) & ". Source XML retained. Open the report to review mapped and remaining blocks.", False)
    Forms!frmPDFImport!cmdOpenResult.Enabled = True
    Forms!frmPDFImport!lblStatus.Caption = "Import complete. Apply is disabled; choose another source for a new import."
    UXApplyPDFImport = True: Exit Function
Failed:
    pdfSignature = ""
    Call ImportState("frmPDFImport", "2. Preview", Err.Description & " Preview again before retrying.", False)
End Function
Private Function UXPDFSignature() As String
    With Forms!frmPDFImport
        UXPDFSignature = PDFImportStamp() & "|" & "20" & "|" & Nz(!cboIdentifier, "") & "|" & Nz(!cboTarget, 0) & "|" & Nz(!cboMergeMode, "Fill empty fields")
    End With
End Function
Public Function UXPDFDependencyCheck() As Boolean
    Forms!frmPDFImport!lblStatus.Caption = PDFExtractionStatus(PDFImportPath())
    UXPDFDependencyCheck = True
End Function
Public Function UXNarrativeInputChanged() As Boolean
    narrativeSource = 0: narrativeHash = "": narrativeCreated = 0
    If CurrentProject.AllForms("frmImportReview").IsLoaded Then
        Forms!frmImportReview!cmdOpenResult.Enabled = False
        Call ImportState("frmImportReview", "2. Preview", "Select a source member and preview narrative mapping. No import performed.", False)
    End If
    UXNarrativeInputChanged = True
End Function
Public Function UXPreviewNarrativeImport() As Boolean
    Dim values As Object, id As Long, text As String
    On Error GoTo Failed
    narrativeSource = 0: narrativeCreated = 0
    Forms!frmImportReview!cmdPreview.SetFocus
    Forms!frmImportReview!cmdOpenResult.Enabled = False
    id = Nz(Forms!frmImportReview!lstImportReports, 0)
    Set values = NarrativeImportValues(id)
    narrativeMember = Nz(Forms!frmImportReview!txtImportMember, "")
    If Len(narrativeMember) > 120 Then Err.Raise vbObjectError + 411, , "Local member display name exceeds 120 characters."
    text = "NEW " & values("FormKind") & " DRAFT - source member " & ImportValue(values("SourceName")) & vbCrLf & "Narrative-only mapping. Official name, dates, traits, duties, senior blocks and folder links are unmapped. Full source values are retained in provenance." & vbCrLf & vbCrLf & UXFieldCaption("OfficialComments", values("FormKind")) & vbCrLf & "Before: [new draft / no prior value]" & vbCrLf & "After: " & ImportValue(values("OfficialComments")) & vbCrLf & vbCrLf & "Local member display name (separate from official name)" & vbCrLf & "Before: [new draft / no prior value]" & vbCrLf & "After: " & ImportValue(narrativeMember)
    Forms!frmImportReview!cmdPreview.SetFocus
    Forms!frmImportReview!uxPreview.Visible = True
    Forms!frmImportReview!lstImportFields.Visible = False
    Forms!frmImportReview!txtFieldValue.Visible = False
    Forms!frmImportReview!lblFieldColumns.Visible = False
    Forms!frmImportReview!uxPreview = text
    narrativeSource = id: narrativeHash = values("SourceHash")
    Call ImportState("frmImportReview", "3. Apply", "Preview ready. Apply creates a partial draft; existing reports are untouched.", True)
    UXPreviewNarrativeImport = True: Exit Function
Failed:
    Call ImportState("frmImportReview", "2. Preview", Err.Description, False)
End Function
Public Function UXApplyNarrativeImport() As Boolean
    On Error GoTo Failed
    If narrativeSource = 0 Or narrativeSource <> Nz(Forms!frmImportReview!lstImportReports, 0) Or narrativeMember <> Nz(Forms!frmImportReview!txtImportMember, "") Then Err.Raise vbObjectError + 4902, , "Preview the selected source member and current display name before Apply."
    narrativeCreated = ImportReviewedNarrative(narrativeSource, narrativeMember, 0, 0, narrativeHash)
    narrativeSource = 0: narrativeHash = ""
    Call ImportState("frmImportReview", "4. Results", "Created " & UXMemberLabel(narrativeCreated) & ". Narrative-only import; source provenance retained. Open the draft to complete its official fields.", False)
    Forms!frmImportReview!cmdOpenResult.Enabled = True
    Forms!frmImportReview!lblImportResult.Caption = "Import complete. Apply is disabled; preview a different source report for another import."
    If CurrentProject.AllForms("frmLibrary").IsLoaded Then Call RefreshLibrary
    UXApplyNarrativeImport = True: Exit Function
Failed:
    narrativeSource = 0: narrativeHash = ""
    Call ImportState("frmImportReview", "2. Preview", Err.Description & " Preview again before retrying.", False)
End Function
Public Function UXOpenImportedReport(ByVal sourceKind As String) As Boolean
    Dim id As Long
    If sourceKind = "PDF" Then id = pdfCreated Else id = narrativeCreated
    If id = 0 Then Exit Function
    If Not UXLeaveEditor() Then Exit Function
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    UXOpenImportedReport = True
End Function
Public Function UXNarrativeInspect() As Boolean
    Dim inspect As Boolean
    With Forms!frmImportReview
        !cmdInspect.SetFocus: inspect = Not !lstImportFields.Visible
        !uxPreview.Visible = Not inspect: !lstImportFields.Visible = inspect
        !txtFieldValue.Visible = inspect: !lblFieldColumns.Visible = inspect
    End With
    UXNarrativeInspect = True
End Function
Public Function UXLeaveImportReview() As Boolean
    Call CloseReviewFile
    DoCmd.Close acForm, "frmImportReview", acSaveNo
    UXLeaveImportReview = UXNavigate("frmLibrary")
End Function
