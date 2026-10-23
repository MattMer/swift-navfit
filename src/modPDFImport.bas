Attribute VB_Name = "modPDFImport"
Option Compare Binary
Option Explicit
Private Type PDFPreviewContext
    Preview As Object
    XML As String
    Path As String
    Hash As String
    Changes As Object
    Target As Long
    Revision As Long
    Choice As String
    StageGuid As String
    OwnerProfileID As Long
End Type
Private pdfContext As PDFPreviewContext








Public Function PreviewStructuredPDF(ByVal path As String) As Object
    Dim packets As Object, packet() As Byte, dom As Object, nodes As Object, result As Object, key As Variant, templateHash As String, kind As String, value As Variant
    Dim failure As Long, message As String, row As Variant, parts As Variant
    On Error GoTo Failed
    Call PDFClearPreview(False)
    Set packets = PDFLocalPackets(path)
    pdfContext.Hash = packets("sourceHash")
    packet = packets("template"): templateHash = PDFBytesHash(packet)
    kind = PDFTemplateKind(templateHash)
    packet = packets("datasets"): pdfContext.XML = PDFPacketText(packet)
    If Len(pdfContext.XML) > 4000000 Then Err.Raise vbObjectError + 4703, , "PDF data exceeds the import size limit."
    Set dom = CreateObject("MSXML2.DOMDocument.6.0"): dom.async = False: dom.resolveExternals = False: dom.preserveWhiteSpace = True: dom.setProperty "ProhibitDTD", True
    If Not dom.LoadXML(pdfContext.XML) Then Err.Raise vbObjectError + 4704, , "Invalid structured PDF data."
    Set result = CreateObject("Scripting.Dictionary"): result.Add "FormKind", kind: result.Add "TemplateHash", templateHash
    For Each row In PDFMapRows(kind)
        parts = Split(CStr(row), "|"): key = parts(1)
        If Not result.Exists(CStr(key)) Then
        Set nodes = dom.selectNodes("//*[local-name()='" & key & "']")
        If nodes.length > 1 Then Err.Raise vbObjectError + 4705, , "Repeated PDF fields require manual review."
        If nodes.length = 0 Then value = Null Else value = CStr(nodes.Item(0).Text)
        result.Add CStr(key), value
        End If
    Next
    If SourceFileHash(path) <> pdfContext.Hash Then Err.Raise vbObjectError + 4706, , "PDF changed during preview."
    pdfContext.Path = path: pdfContext.OwnerProfileID = WizardPreviewOwner(): Set pdfContext.Preview = result: Set PreviewStructuredPDF = result: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    Set pdfContext.Preview = Nothing: pdfContext.XML = "": pdfContext.Path = "": pdfContext.Hash = ""
    On Error GoTo 0
    Err.Raise failure, "PreviewStructuredPDF", message
End Function
Private Function PDFDate(ByVal value As Variant, ByVal century As String) As Variant
    Dim text As String
    If IsNull(value) Then PDFDate = Null: Exit Function
    text = Trim$(CStr(value))
    If text = "" Then PDFDate = Null: Exit Function
    If Len(text) = 6 Then
        text = "20" & Left$(text, 2) & "-" & Mid$(text, 3, 2) & "-" & Right$(text, 2)
    ElseIf Len(text) = 8 Then
        text = Left$(text, 4) & "-" & Mid$(text, 5, 2) & "-" & Right$(text, 2)
    End If
    If Not IsValidIsoDate(text) Then Err.Raise vbObjectError + 4711, , "PDF date needs manual review: " & text
    PDFDate = DateSerial(CLng(Left$(text, 4)), CLng(Mid$(text, 6, 2)), CLng(Right$(text, 2)))
End Function
Public Function PDFMappedValues(ByVal century As String, ByVal identifierKind As String) As Object
    PDFRequireOwner
    Dim changes As Object, row As Variant, parts As Variant, key As String, value As Variant
    If pdfContext.Preview Is Nothing Then Err.Raise vbObjectError + 4712, , "Choose a supported PDF first."
    If SourceFileHash(pdfContext.Path) <> pdfContext.Hash Then Err.Raise vbObjectError + 4706, , "PDF changed after preview. Choose it again."
    If identifierKind <> "SSN" And identifierKind <> "DoDID" And identifierKind <> "Retain only" Then Err.Raise vbObjectError + 4714, , "Choose how to interpret the identifier."
    Set changes = CreateObject("Scripting.Dictionary")
    For Each row In PDFMapRows(pdfContext.Preview("FormKind"))
        parts = Split(CStr(row), "|"): key = parts(0): value = pdfContext.Preview(parts(1))
        ' Missing and empty source fields never erase a saved draft value.
        If Not IsNull(value) Then
            If Len(Trim$(CStr(value))) > 0 Then
                value = PDFConvertValue(CStr(value), CStr(parts(2)), key, century)
                If key = "MemberSSN" Then
                    If identifierKind = "DoDID" Then key = "MemberDoDID"
                    If identifierKind = "Retain only" Then key = ""
                End If
                If Len(key) > 0 And Not IsNull(value) Then changes.Add key, value
            End If
        End If
    Next
    Set PDFMappedValues = changes
End Function
Private Function PDFConvertValue(ByVal value As String, ByVal conversion As String, ByVal key As String, ByVal century As String) As Variant
    Dim mark As String, choices As Variant, index As Long
    mark = LCase$(Trim$(value))
    Select Case conversion
        Case "text": PDFConvertValue = value
        Case "date": PDFConvertValue = PDFDate(value, century)
        Case "date-text"
            If key = "CounselingDate" And IsCounselingStatus(value) Then PDFConvertValue = value Else PDFConvertValue = Format$(PDFDate(value, century), "yyyy-mm-dd")
        Case "flag", "observation"
            If mark <> "0" And mark <> "1" And mark <> "2" Then Err.Raise vbObjectError + 4721, , "Unrecognized PDF checkbox value for " & key & ": " & value
            If mark = "2" Then PDFConvertValue = Null: Exit Function
            If conversion = "flag" Then PDFConvertValue = IIf(mark = "1", "Yes", "No") Else PDFConvertValue = IIf(mark = "1", "NOB", "Observed")
        Case "component"
            If mark <> "act" And mark <> "tar" And mark <> "inact" And mark <> "at" Then Err.Raise vbObjectError + 4721, , "Unrecognized PDF component: " & value
            choices = Array("ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW")
            For index = 0 To 3
                If choices(index) = key Then PDFConvertValue = IIf(Split("act|tar|inact|at", "|")(index) = mark, "Yes", "No")
            Next
        Case "trait", "chief-trait"
            If mark = "nob" Or (conversion = "chief-trait" And mark = "0") Then PDFConvertValue = "NOB": Exit Function
            If Len(mark) = 1 And mark >= "1" And mark <= "5" Then PDFConvertValue = mark & ".0": Exit Function
            If mark = "1.0" Or mark = "2.0" Or mark = "3.0" Or mark = "4.0" Or mark = "5.0" Then PDFConvertValue = mark: Exit Function
            Err.Raise vbObjectError + 4721, , "Unrecognized PDF trait for " & key & ": " & value
        Case "promotion-status"
            Select Case mark
                Case "1": PDFConvertValue = "REGULAR"
                Case "2": PDFConvertValue = "FROCKED"
                Case "3": PDFConvertValue = "SELECTED"
                Case "4": PDFConvertValue = "SPOT"
                Case "regular", "frocked", "selected", "spot": PDFConvertValue = UCase$(mark)
                Case Else: Err.Raise vbObjectError + 4721, , "Unrecognized PDF promotion status."
            End Select
        Case "recommendation"
            Select Case mark
                Case "nob": PDFConvertValue = "NOB"
                Case "sp": PDFConvertValue = "Significant Problems"
                Case "pg": PDFConvertValue = "Progressing"
                Case "pm": PDFConvertValue = "Promotable"
                Case "mp": PDFConvertValue = "Must Promote"
                Case "ep": PDFConvertValue = "Early Promote"
                Case Else: Err.Raise vbObjectError + 4721, , "Unrecognized PDF promotion recommendation."
            End Select
        Case Else: Err.Raise vbObjectError + 4721, , "Unknown PDF conversion."
    End Select
End Function
Public Function PDFClearPreview(Optional ByVal cancelSaved As Boolean = True) As Boolean
    If cancelSaved Then Call WizardPreviewCancel("PDFImport")
    Set pdfContext.Preview = Nothing: pdfContext.XML = "": pdfContext.Path = "": pdfContext.Hash = ""
    Set pdfContext.Changes = Nothing: pdfContext.Target = 0: pdfContext.Revision = 0: pdfContext.Choice = ""
    pdfContext.StageGuid = ""
    pdfContext.OwnerProfileID = 0
    PDFClearPreview = True
End Function
Public Function PDFImportKind() As String
    If Not pdfContext.Preview Is Nothing Then PDFImportKind = pdfContext.Preview("FormKind")
End Function
Public Function PDFImportPath() As String
    PDFImportPath = pdfContext.Path
End Function
Public Function PDFImportStamp() As String
    PDFImportStamp = pdfContext.Hash & "|" & pdfContext.Path
End Function
Public Function ImportPDFPreview(ByVal century As String, ByVal identifierKind As String, Optional ByVal targetID As Long = 0, Optional ByVal mode As String = "Fill empty fields") As Long
    Dim id As Long, changes As Object, receipt As DAO.Recordset, raw As DAO.Recordset, receiptID As Long, w As DAO.Workspace, failure As Long, message As String, db As DAO.Database, stage As String, expected As Long, operationID As Long
    If pdfContext.Preview Is Nothing Then Err.Raise vbObjectError + 4712, , "Preview a supported PDF first."
    PDFRequireOwner
    If pdfContext.StageGuid <> "" Then WizardPreviewCheck "PDFImport", pdfContext.StageGuid
    If SourceFileHash(pdfContext.Path) <> pdfContext.Hash Then Err.Raise vbObjectError + 4706, , "PDF changed after preview."
    If DCount("*", "ImportReceipt", "SourceHash=" & QuoteSQL(pdfContext.Hash) & " AND AdapterVersion Like 'pdf-xfa-v*'") > 0 Then Err.Raise vbObjectError + 4713, , "This exact PDF has already been imported."
    If identifierKind <> "SSN" And identifierKind <> "DoDID" And identifierKind <> "Retain only" Then Err.Raise vbObjectError + 4714, , "Choose how to interpret the combined identifier field."
    If targetID > 0 Then
        If pdfContext.Changes Is Nothing Then Err.Raise vbObjectError + 4722, , "Preview the merge before Apply."
        If pdfContext.Target <> targetID Or pdfContext.Choice <> century & "|" & identifierKind & "|" & mode Then Err.Raise vbObjectError + 4722, , "Merge choices changed. Preview again."
        Call PDFCheckTarget(targetID)
        If CurrentRevision(targetID) <> pdfContext.Revision Then Err.Raise vbObjectError + 4722, , "Target draft changed after preview. Preview again."
        Set changes = pdfContext.Changes: expected = pdfContext.Revision
    Else
        Set changes = PDFMappedValues(century, identifierKind)
    End If
    If changes.Count = 0 Then Err.Raise vbObjectError + 4722, , "No applicable nonblank PDF values to apply."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    operationID = BeginOperation("PDF import", pdfContext.Path)
    stage = "Create draft"
    If targetID = 0 Then
        'Create the directory link from extracted identity, never a placeholder.
        Dim identity As Object, identityKey As Variant
        Set identity = CreateObject("Scripting.Dictionary")
        For Each identityKey In Array("OfficialName", "RankRate", "Designator", "UnitUIC", "ShipStation", "MemberSSN", "MemberDoDID", "ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW", "DateReported")
            identity.Add CStr(identityKey), Null
            If changes.Exists(CStr(identityKey)) Then identity(identityKey) = changes(identityKey)
        Next
        id = CommitDraft(0, pdfContext.Preview("FormKind"), "Imported PDF draft", Nz(identity("OfficialName"), "Unnamed member"), Null, Null, "Unmapped PDF fields are retained as source XML and require review.", "PDF import", administration:=FieldsSubset(identity, AdministrationKeys()), memberDetails:=FieldsSubset(identity, MemberDetailsKeys()))
        expected = CurrentRevision(id)
        If changes.Exists("OfficialName") Then changes.Add "MemberName", Left$(CStr(changes("OfficialName")), 120)
    Else
        id = targetID
    End If
    stage = "Commit PDF fields"
    Call CommitSavedFields(id, expected, changes, "Reviewed structured PDF import")
    If CurrentRevision(id) <> expected Then AuditCorrelateRevision CurrentRevision(id), operationID
    stage = "Retain receipt": Set db = CurrentDb
    Set receipt = db.OpenRecordset("ImportReceipt", dbOpenDynaset)
    receipt.AddNew: receipt!ReceiptGuid = NewIdentity(): receipt!SourceHash = pdfContext.Hash: receipt!SourcePath = pdfContext.Path: receipt!ExternalReportID = 0: receipt!ReportID = id
    receipt!AuthorProfileID = MajorAuthor(): receipt!ImportedAt = Now(): receipt!AdapterVersion = "pdf-xfa-v3-local": receipt!ValueEncoding = "XFA XML; " & IIf(targetID = 0, "new draft", mode) & "; baseline " & expected: receipt!AppliedRevisionID = CurrentRevision(id): receipt!IsUndone = False
    If targetID > 0 Then receipt!BeforeRevisionID = expected
    receipt!OperationID = operationID
    receipt.Update: receipt.Bookmark = receipt.LastModified: receiptID = receipt!ImportReceiptID: receipt.Close
    stage = "Retain XML"
    Set raw = db.OpenRecordset("ImportRawValue", dbOpenDynaset)
    raw.AddNew: raw!ImportReceiptID = receiptID: raw!SourceTable = "PDF-XFA": raw!SourceRow = 0: raw!FieldOrdinal = 0: raw!FieldKey = "datasets": raw!DAOType = dbMemo: raw!RawValue = pdfContext.XML: raw!IsNullValue = False: raw.Update
    raw.AddNew: raw!ImportReceiptID = receiptID: raw!SourceTable = "PDF-XFA": raw!SourceRow = 0: raw!FieldOrdinal = 1: raw!FieldKey = "template-hash": raw!DAOType = dbText: raw!RawValue = pdfContext.Preview("TemplateHash"): raw!IsNullValue = False: raw.Update: raw.Close
    If SourceFileHash(pdfContext.Path) <> pdfContext.Hash Then Err.Raise vbObjectError + 4706, , "PDF changed during import."
    Call FinishOperation(operationID, "Completed", "PDF values applied to report " & id & "; source XML retained.")
    If pdfContext.StageGuid <> "" Then Call WizardPreviewCancel("PDFImport", pdfContext.StageGuid)
    w.CommitTrans: ImportPDFPreview = id: Call PDFClearPreview(False): Exit Function
Failed:
    failure = Err.Number: message = stage & ": " & Err.Description
    On Error Resume Next
    w.Rollback
    If Not receipt Is Nothing Then receipt.Close
    If Not raw Is Nothing Then raw.Close
    On Error GoTo 0
    Err.Raise failure, "ImportPDFPreview", message
End Function
Public Sub PDFCheckpointChoices(ByVal century As String, ByVal identifierKind As String, ByVal target As Long, ByVal mode As String)
    Dim choices As Object, revision As Long
    If pdfContext.Preview Is Nothing Then Err.Raise vbObjectError + 4712, , "Choose a PDF before previewing."
    PDFRequireOwner
    Set choices = CreateObject("Scripting.Dictionary")
    choices.Add "Century", century: choices.Add "Identifier", identifierKind: choices.Add "Mode", mode
    If target > 0 Then revision = pdfContext.Revision
    pdfContext.StageGuid = WizardPreviewSave("PDFImport", pdfContext.Path, pdfContext.Hash, "", "", target, revision, choices, pdfContext.StageGuid)
End Sub
Private Sub PDFRequireOwner()
    If pdfContext.Preview Is Nothing Then Err.Raise vbObjectError + 4712, , "Preview ownership was reset. Resume the saved PDF preview first."
    If pdfContext.OwnerProfileID <> WizardPreviewOwner() Then Err.Raise vbObjectError + 9203, , "This PDF preview belongs to another contributor. Resume or preview again."
End Sub
Public Function PDFResumePreview() As Object
    Dim saved As Object, preview As Object, changes As Object
    Set saved = WizardPreviewLoad("PDFImport")
    WizardPreviewCheck "PDFImport", saved("StageGuid")
    Set preview = PreviewStructuredPDF(saved("SourcePath"))
    If Not IsNull(saved("TargetReportID")) Then
        Set changes = PDFPrepareMerge(saved("Choices")("Century"), saved("Choices")("Identifier"), saved("TargetReportID"), saved("Choices")("Mode"))
    End If
    pdfContext.StageGuid = saved("StageGuid")
    Set PDFResumePreview = saved
End Function
Private Sub PDFCheckTarget(ByVal id As Long)
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT FormKind,IsArchived,WorkflowState FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 4723, , "Target draft does not exist."
    If r!FormKind <> PDFImportKind() Or r!IsArchived Or r!WorkflowState <> "Draft" Then r.Close: Err.Raise vbObjectError + 4723, , "Choose an editable draft of the same report type."
    r.Close
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Nz(Forms!frmEditor!txtReportID, 0) = id Then Err.Raise vbObjectError + 4723, , "Close the target report editor before merging."
    End If
    If DCount("*", "RecoveryDraft", "ReportID=" & id) > 0 Then Err.Raise vbObjectError + 4723, , "Resolve the target draft's pending recovery before merging."
End Sub
Public Function PDFPrepareMerge(ByVal century As String, ByVal identifierKind As String, ByVal id As Long, ByVal mode As String) As Object
    Dim incoming As Object, before As Object, result As Object, key As Variant
    Set pdfContext.Changes = Nothing: pdfContext.Target = 0: pdfContext.Revision = 0
    If mode <> "Fill empty fields" And mode <> "Replace mapped values" Then Err.Raise vbObjectError + 4723, , "Choose a merge mode."
    Call PDFCheckTarget(id)
    pdfContext.Revision = CurrentRevision(id)
    Set incoming = PDFMappedValues(century, identifierKind): Set before = SavedFields(id)
    If CurrentRevision(id) <> pdfContext.Revision Then Err.Raise vbObjectError + 4722, , "Target draft changed during preview. Preview again."
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In incoming.Keys
        If (mode = "Replace mapped values" Or Len(Nz(before(key), "")) = 0) And Not SameFieldValue(before(key), incoming(key)) Then result.Add CStr(key), incoming(key)
    Next
    pdfContext.Target = id: pdfContext.Choice = century & "|" & identifierKind & "|" & mode
    Set pdfContext.Changes = result: Set PDFPrepareMerge = result
End Function
Private Function PDFTemplateKind(ByVal templateHash As String) As String
    Select Case templateHash
        Case "54338AECE7EDDA4B2BE41389262219AACE3DB5EAAD59634D31C062E1AC4EE134": PDFTemplateKind = "EVAL"
        Case "2F554C052290D1820604B6341438C861EAB251D89E4B2D7B2BD3A108C679532E": PDFTemplateKind = "CHIEFEVAL"
        Case "41A9E82DA6016B25C3BB77604BEC66B453BF78B637659D008E3E8B897726CE91": PDFTemplateKind = "FITREP"
        Case Else: Err.Raise vbObjectError + 4702, , "Unsupported PDF form revision. Choose a supported May 2025 NAVPERS fillable form; flattened/scanned forms need manual entry."
    End Select
End Function
Public Function PDFExtractionStatus(Optional ByVal testPath As String = "") As String
    Dim packets As Object, packet() As Byte, dom As Object, kind As String
    On Error GoTo Failed
    If Len(testPath) = 0 Then
        PDFExtractionStatus = "Local PDF extraction available. Choose a supported May 2025 fillable NAVPERS PDF to check it. No Acrobat or internet connection required."
        Exit Function
    End If
    ' Probe independently: do not clear an existing preview or prepared merge.
    Set packets = PDFLocalPackets(testPath)
    packet = packets("template"): kind = PDFTemplateKind(PDFBytesHash(packet))
    packet = packets("datasets")
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False: dom.resolveExternals = False: dom.setProperty "ProhibitDTD", True
    If Not dom.LoadXML(PDFPacketText(packet)) Then Err.Raise 5, , "Invalid structured PDF data."
    PDFExtractionStatus = "PASS: local read-only " & kind & " XFA extraction; no Acrobat required."
    Exit Function
Failed:
    PDFExtractionStatus = "PDF extraction unavailable for this file: " & Err.Description
End Function
Public Function ChoosePDFPreview() As Boolean
    Dim dialog As Object, result As Object, key As Variant, text As String
    On Error GoTo Failed
    Set dialog = Application.FileDialog(3): dialog.Title = "Preview a completed NAVPERS PDF": dialog.AllowMultiSelect = False
    dialog.Filters.Clear: dialog.Filters.Add "PDF forms", "*.pdf"
    If dialog.Show <> -1 Then Exit Function
    ChoosePDFPreview = UXChoosePDFFile(dialog.SelectedItems(1)): Exit Function
Failed:
    Forms!frmPDFImport!lblStatus.Caption = Err.Description
End Function
Public Function ImportPDFUI() As Boolean
    ImportPDFUI = UXApplyPDFImport()
End Function
