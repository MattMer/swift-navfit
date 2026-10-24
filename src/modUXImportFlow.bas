Attribute VB_Name = "modUXImportFlow"
Option Compare Binary
Option Explicit
Private Type NarrativePreviewContext
    SourceID As Long
    Hash As String
    Member As String
    Created As Long
    OwnerProfileID As Long
End Type
Private narrativePreview As NarrativePreviewContext
Private pdfSignature As String
Private pdfCreated As Long




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
    Call PDFClearPreview(False)
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
    If Nz(Forms!frmPDFImport!cboIdentifier, "") = "DoDID" And Not Nz(Forms!frmPDFImport!chkLegacy, False) Then Err.Raise vbObjectError + 4903, "modUXImportFlow.UXPreviewPDFImport", "Enable legacy compatibility explicitly before mapping DoDID."
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
    If changes.Count = 0 Then Err.Raise vbObjectError + 4904, "modUXImportFlow.UXPreviewPDFImport", "No changed values to apply. Choose another merge mode or target."
    Forms!frmPDFImport!txtPreview = text
    PDFCheckpointChoices "20", Nz(Forms!frmPDFImport!cboIdentifier, ""), target, mode
    pdfSignature = UXPDFSignature()
    Call ImportState("frmPDFImport", "3. Apply", CStr(changes.Count) & " values ready. Review full before/after values; Apply retains source provenance.", True)
    UXPreviewPDFImport = True: Exit Function
Failed:
    Call ImportState("frmPDFImport", "2. Preview", Err.Description, False)
End Function
Public Function UXResumePDFImport() As Boolean
    Dim saved As Object
    On Error GoTo Failed
    pdfSignature = "": Forms!frmPDFImport!cmdImport.Enabled = False
    Set saved = PDFResumePreview()
    With Forms!frmPDFImport
        !chkLegacy = (saved("Choices")("Identifier") = "DoDID")
        Call UXPDFLegacyChanged
        !cboIdentifier = saved("Choices")("Identifier"): !cboMergeMode = saved("Choices")("Mode")
        !cboTarget = saved("TargetReportID")
    End With
    UXResumePDFImport = UXPreviewPDFImport(): Exit Function
Failed:
    Call ImportState("frmPDFImport", "Resume unavailable", Err.Description & " Saved choices are retained.", False)
End Function
Public Function UXCancelPDFImport() As Boolean
    On Error GoTo Failed
    UXCancelPDFImport = PDFClearPreview()
    On Error GoTo 0
    pdfSignature = "": Forms!frmPDFImport!txtPreview = ""
    Call ImportState("frmPDFImport", "Cancelled", "Saved preview removed. No report values changed.", False)
    Exit Function
Failed:
    Call ImportState("frmPDFImport", "Cancel unavailable", Err.Description & " Saved choices are retained. Resume before cancelling.", False)
End Function
Public Function UXApplyPDFImport() As Boolean
    Dim outcome As String, message As String
    outcome = "Not started"
    On Error GoTo Failed
    If Len(pdfSignature) = 0 Or pdfSignature <> UXPDFSignature() Then Err.Raise vbObjectError + 4901, "modUXImportFlow.UXApplyPDFImport", "Preview the current mapping before Apply."
    pdfCreated = ImportPDFPreview("20", Nz(Forms!frmPDFImport!cboIdentifier, ""), Nz(Forms!frmPDFImport!cboTarget, 0), Nz(Forms!frmPDFImport!cboMergeMode, "Fill empty fields"), outcome)
    pdfSignature = ""
    Call ImportState("frmPDFImport", "4. Results", "Imported into " & UXMemberLabel(pdfCreated) & ". Source XML retained. Open the report to review mapped and remaining blocks.", False)
    Forms!frmPDFImport!cmdOpenResult.Enabled = True
    Forms!frmPDFImport!lblStatus.Caption = "Import complete. Apply is disabled; choose another source for a new import."
    UXApplyPDFImport = True: Exit Function
Failed:
    message = Err.Description
    pdfSignature = ""
    Select Case outcome
    Case "Committed"
        Call PDFClearPreview(False)
        Call ImportState("frmPDFImport", "4. Results", "Import data was committed; finalization failed. Do not re-import. Close and reopen this workspace. " & message, False)
    Case "Not started", "Rolled back"
        Call ImportState("frmPDFImport", "2. Preview", "Import " & LCase$(outcome) & ". " & message & " Preview again before retrying.", False)
    Case Else
        Call PDFClearPreview(False)
        Call ImportState("frmPDFImport", "Inspect saved data", "Import outcome could not be confirmed. Close and reopen; inspect saved data before retrying. " & message, False)
    End Select
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
    narrativePreview.SourceID = 0: narrativePreview.Hash = "": narrativePreview.Created = 0
    If CurrentProject.AllForms("frmImportReview").IsLoaded Then
        Forms!frmImportReview!cmdOpenResult.Enabled = False
        Call ImportState("frmImportReview", "2. Preview", "Select a source member and preview narrative mapping. No import performed.", False)
    End If
    UXNarrativeInputChanged = True
End Function
Public Function UXPreviewNarrativeImport() As Boolean
    Dim values As Object, id As Long, text As String
    On Error GoTo Failed
    narrativePreview.SourceID = 0: narrativePreview.Created = 0
    Forms!frmImportReview!cmdPreview.SetFocus
    Forms!frmImportReview!cmdOpenResult.Enabled = False
    id = Nz(Forms!frmImportReview!lstImportReports, 0)
    Set values = NarrativeImportValues(id)
    narrativePreview.Member = Nz(Forms!frmImportReview!txtImportMember, "")
    If Len(narrativePreview.Member) > 120 Then Err.Raise vbObjectError + 411, "modUXImportFlow.UXPreviewNarrativeImport", "Local member display name exceeds 120 characters."
    text = "NEW " & values("FormKind") & " DRAFT - source member " & ImportValue(values("SourceName")) & vbCrLf & "Narrative-only mapping. Official name, dates, traits, duties, senior blocks and folder links are unmapped. Full source values are retained in provenance." & vbCrLf & vbCrLf & UXFieldCaption("OfficialComments", values("FormKind")) & vbCrLf & "Before: [new draft / no prior value]" & vbCrLf & "After: " & ImportValue(values("OfficialComments")) & vbCrLf & vbCrLf & "Local member display name (separate from official name)" & vbCrLf & "Before: [new draft / no prior value]" & vbCrLf & "After: " & ImportValue(narrativePreview.Member)
    Forms!frmImportReview!cmdPreview.SetFocus
    Forms!frmImportReview!uxPreview.Visible = True
    Forms!frmImportReview!lstImportFields.Visible = False
    Forms!frmImportReview!txtFieldValue.Visible = False
    Forms!frmImportReview!lblFieldColumns.Visible = False
    Forms!frmImportReview!uxPreview = text
    narrativePreview.SourceID = id: narrativePreview.Hash = values("SourceHash")
    narrativePreview.OwnerProfileID = ActiveContributorID()
    Call ImportState("frmImportReview", "3. Apply", "Preview ready. Apply creates a partial draft; existing reports are untouched.", True)
    UXPreviewNarrativeImport = True: Exit Function
Failed:
    Call ImportState("frmImportReview", "2. Preview", Err.Description, False)
End Function
Public Function UXApplyNarrativeImport() As Boolean
    On Error GoTo Failed
    If narrativePreview.OwnerProfileID <> ActiveContributorID() Then Err.Raise vbObjectError + 9203, "modUXImportFlow.UXApplyNarrativeImport", "Preview this import as the current contributor before Apply."
    If narrativePreview.SourceID = 0 Or narrativePreview.SourceID <> Nz(Forms!frmImportReview!lstImportReports, 0) Or narrativePreview.Member <> Nz(Forms!frmImportReview!txtImportMember, "") Then Err.Raise vbObjectError + 4902, "modUXImportFlow.UXApplyNarrativeImport", "Preview the selected source member and current display name before Apply."
    narrativePreview.Created = ImportReviewedNarrative(narrativePreview.SourceID, narrativePreview.Member, 0, 0, narrativePreview.Hash)
    narrativePreview.SourceID = 0: narrativePreview.Hash = ""
    Call ImportState("frmImportReview", "4. Results", "Created " & UXMemberLabel(narrativePreview.Created) & ". Narrative-only import; source provenance retained. Open the draft to complete its official fields.", False)
    Forms!frmImportReview!cmdOpenResult.Enabled = True
    Forms!frmImportReview!lblImportResult.Caption = "Import complete. Apply is disabled; preview a different source report for another import."
    If CurrentProject.AllForms("frmLibrary").IsLoaded Then Call RefreshLibrary
    UXApplyNarrativeImport = True: Exit Function
Failed:
    narrativePreview.SourceID = 0: narrativePreview.Hash = ""
    Call ImportState("frmImportReview", "2. Preview", Err.Description & " Preview again before retrying.", False)
End Function
Public Function UXOpenImportedReport(ByVal sourceKind As String) As Boolean
    Dim id As Long
    If sourceKind = "PDF" Then id = pdfCreated Else id = narrativePreview.Created
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
