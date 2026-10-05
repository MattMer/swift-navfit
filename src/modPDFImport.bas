Attribute VB_Name = "modPDFImport"
Option Compare Binary
Option Explicit
Private pdfPreview As Object
Private pdfXML As String
Private pdfPath As String
Private pdfHash As String
Private pdfChanges As Object
Private pdfTarget As Long
Private pdfRevision As Long
Private pdfChoice As String
Public Function PreviewStructuredPDF(ByVal path As String) As Object
    Dim p As Object, j As Object, x As Object, dom As Object, nodes As Object, result As Object, fields As Variant, key As Variant, templateHash As String, kind As String, value As Variant
    Dim failure As Long, message As String, row As Variant, parts As Variant
    On Error GoTo Failed
    Call PDFClearPreview
    If FileLen(path) > 30000000 Then Err.Raise vbObjectError + 4700, , "PDF exceeds the import size limit."
    pdfHash = SourceFileHash(path)
    Set p = CreateObject("AcroExch.PDDoc")
    If Not p.Open(path) Then Err.Raise vbObjectError + 4701, , "Acrobat could not open this PDF."
    Set j = p.GetJSObject(): Set x = j.xfa
    templateHash = SourceTextHash(CStr(x.template.saveXML()))
    Select Case templateHash
        Case "B0184F00EC5C2F99CD6CE96FD4FD1A80DD142A3AA62C9BE53D4E51ECF8DAF076": kind = "EVAL"
        Case "E1A0A124D4F927875C9920FE1B4004B6A90C0E69FA059B1D5F835E8AB1640EF4": kind = "CHIEFEVAL"
        Case "4C006AC4201AC12472F8CA8D830C3010E4502D0685BFB93B9F54BEE7516C08D3": kind = "FITREP"
        Case Else: Err.Raise vbObjectError + 4702, , "Unsupported PDF form revision. Flattened forms cannot be imported as structured data."
    End Select
    pdfXML = CStr(x.datasets.saveXML())
    p.Close: Set p = Nothing
    If Len(pdfXML) > 4000000 Then Err.Raise vbObjectError + 4703, , "PDF data exceeds the import size limit."
    Set dom = CreateObject("MSXML2.DOMDocument.6.0"): dom.async = False: dom.resolveExternals = False: dom.preserveWhiteSpace = True: dom.setProperty "ProhibitDTD", True
    If Not dom.LoadXML(pdfXML) Then Err.Raise vbObjectError + 4704, , "Invalid structured PDF data."
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
    If SourceFileHash(path) <> pdfHash Then Err.Raise vbObjectError + 4706, , "PDF changed during preview."
    pdfPath = path: Set pdfPreview = result: Set PreviewStructuredPDF = result: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not p Is Nothing Then p.Close
    Set pdfPreview = Nothing: pdfXML = "": pdfPath = "": pdfHash = ""
    On Error GoTo 0
    Err.Raise failure, "PreviewStructuredPDF", message
End Function
Private Function PDFDate(ByVal value As Variant, ByVal century As String) As Variant
    Dim text As String
    If IsNull(value) Then PDFDate = Null: Exit Function
    text = Trim$(CStr(value))
    If text = "" Then PDFDate = Null: Exit Function
    If Len(text) = 6 Then
        If century <> "19" And century <> "20" And century <> "21" Then Err.Raise vbObjectError + 4710, , "Choose a century explicitly for two-digit PDF dates."
        text = century & Left$(text, 2) & "-" & Mid$(text, 3, 2) & "-" & Right$(text, 2)
    ElseIf Len(text) = 8 Then
        text = Left$(text, 4) & "-" & Mid$(text, 5, 2) & "-" & Right$(text, 2)
    End If
    If Not IsValidIsoDate(text) Then Err.Raise vbObjectError + 4711, , "PDF date needs manual review: " & text
    PDFDate = DateSerial(CLng(Left$(text, 4)), CLng(Mid$(text, 6, 2)), CLng(Right$(text, 2)))
End Function
Public Function PDFMappedValues(ByVal century As String, ByVal identifierKind As String) As Object
    Dim changes As Object, row As Variant, parts As Variant, key As String, value As Variant
    If pdfPreview Is Nothing Then Err.Raise vbObjectError + 4712, , "Choose a supported PDF first."
    If SourceFileHash(pdfPath) <> pdfHash Then Err.Raise vbObjectError + 4706, , "PDF changed after preview. Choose it again."
    If identifierKind <> "SSN" And identifierKind <> "DoDID" And identifierKind <> "Retain only" Then Err.Raise vbObjectError + 4714, , "Choose how to interpret the identifier."
    Set changes = CreateObject("Scripting.Dictionary")
    For Each row In PDFMapRows(pdfPreview("FormKind"))
        parts = Split(CStr(row), "|"): key = parts(0): value = pdfPreview(parts(1))
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
        Case "date-text": PDFConvertValue = Format$(PDFDate(value, century), "yyyy-mm-dd")
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
Public Function PDFClearPreview() As Boolean
    Set pdfPreview = Nothing: pdfXML = "": pdfPath = "": pdfHash = ""
    Set pdfChanges = Nothing: pdfTarget = 0: pdfRevision = 0: pdfChoice = ""
    PDFClearPreview = True
End Function
Public Function PDFImportKind() As String
    If Not pdfPreview Is Nothing Then PDFImportKind = pdfPreview("FormKind")
End Function
Public Function PDFImportPath() As String
    PDFImportPath = pdfPath
End Function
Public Function PDFImportStamp() As String
    PDFImportStamp = pdfHash & "|" & pdfPath
End Function
Public Function ImportPDFPreview(ByVal century As String, ByVal identifierKind As String, Optional ByVal targetID As Long = 0, Optional ByVal mode As String = "Fill empty fields") As Long
    Dim id As Long, changes As Object, receipt As DAO.Recordset, raw As DAO.Recordset, receiptID As Long, w As DAO.Workspace, failure As Long, message As String, db As DAO.Database, stage As String, expected As Long, operationID As Long
    If pdfPreview Is Nothing Then Err.Raise vbObjectError + 4712, , "Preview a supported PDF first."
    If SourceFileHash(pdfPath) <> pdfHash Then Err.Raise vbObjectError + 4706, , "PDF changed after preview."
    If DCount("*", "ImportReceipt", "SourceHash=" & QuoteSQL(pdfHash) & " AND AdapterVersion Like 'pdf-xfa-v*'") > 0 Then Err.Raise vbObjectError + 4713, , "This exact PDF has already been imported."
    If identifierKind <> "SSN" And identifierKind <> "DoDID" And identifierKind <> "Retain only" Then Err.Raise vbObjectError + 4714, , "Choose how to interpret the combined identifier field."
    If targetID > 0 Then
        If pdfChanges Is Nothing Then Err.Raise vbObjectError + 4722, , "Preview the merge before Apply."
        If pdfTarget <> targetID Or pdfChoice <> century & "|" & identifierKind & "|" & mode Then Err.Raise vbObjectError + 4722, , "Merge choices changed. Preview again."
        Call PDFCheckTarget(targetID)
        If CurrentRevision(targetID) <> pdfRevision Then Err.Raise vbObjectError + 4722, , "Target draft changed after preview. Preview again."
        Set changes = pdfChanges: expected = pdfRevision
    Else
        Set changes = PDFMappedValues(century, identifierKind)
    End If
    If changes.Count = 0 Then Err.Raise vbObjectError + 4722, , "No applicable nonblank PDF values to apply."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    operationID = BeginOperation("PDF import", pdfPath)
    stage = "Create draft"
    If targetID = 0 Then
        id = CommitDraft(0, pdfPreview("FormKind"), "Imported PDF draft", "Imported PDF member", Null, Null, "Unmapped PDF fields are retained as source XML and require review.", "PDF import")
        expected = CurrentRevision(id)
        If changes.Exists("OfficialName") Then changes.Add "MemberName", Left$(CStr(changes("OfficialName")), 120)
    Else
        id = targetID
    End If
    stage = "Commit PDF fields"
    Call CommitSavedFields(id, expected, changes, "Reviewed structured PDF import")
    stage = "Retain receipt": Set db = CurrentDb
    Set receipt = db.OpenRecordset("ImportReceipt", dbOpenDynaset)
    receipt.AddNew: receipt!ReceiptGuid = NewIdentity(): receipt!SourceHash = pdfHash: receipt!SourcePath = pdfPath: receipt!ExternalReportID = 0: receipt!ReportID = id
    receipt!AuthorProfileID = MajorAuthor(): receipt!ImportedAt = Now(): receipt!AdapterVersion = "pdf-xfa-v2": receipt!ValueEncoding = "XFA XML; " & IIf(targetID = 0, "new draft", mode) & "; baseline " & expected: receipt!AppliedRevisionID = CurrentRevision(id): receipt!IsUndone = False
    If targetID > 0 Then receipt!BeforeRevisionID = expected
    receipt!OperationID = operationID
    receipt.Update: receipt.Bookmark = receipt.LastModified: receiptID = receipt!ImportReceiptID: receipt.Close
    stage = "Retain XML"
    Set raw = db.OpenRecordset("ImportRawValue", dbOpenDynaset)
    raw.AddNew: raw!ImportReceiptID = receiptID: raw!SourceTable = "PDF-XFA": raw!SourceRow = 0: raw!FieldOrdinal = 0: raw!FieldKey = "datasets": raw!DAOType = dbMemo: raw!RawValue = pdfXML: raw!IsNullValue = False: raw.Update
    raw.AddNew: raw!ImportReceiptID = receiptID: raw!SourceTable = "PDF-XFA": raw!SourceRow = 0: raw!FieldOrdinal = 1: raw!FieldKey = "template-hash": raw!DAOType = dbText: raw!RawValue = pdfPreview("TemplateHash"): raw!IsNullValue = False: raw.Update: raw.Close
    If SourceFileHash(pdfPath) <> pdfHash Then Err.Raise vbObjectError + 4706, , "PDF changed during import."
    Call FinishOperation(operationID, "Completed", "PDF values applied to report " & id & "; source XML retained.")
    w.CommitTrans: ImportPDFPreview = id: Call PDFClearPreview: Exit Function
Failed:
    failure = Err.Number: message = stage & ": " & Err.Description
    On Error Resume Next
    w.Rollback
    If Not receipt Is Nothing Then receipt.Close
    If Not raw Is Nothing Then raw.Close
    On Error GoTo 0
    Err.Raise failure, "ImportPDFPreview", message
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
    Set pdfChanges = Nothing: pdfTarget = 0: pdfRevision = 0
    If mode <> "Fill empty fields" And mode <> "Replace mapped values" Then Err.Raise vbObjectError + 4723, , "Choose a merge mode."
    Call PDFCheckTarget(id)
    pdfRevision = CurrentRevision(id)
    Set incoming = PDFMappedValues(century, identifierKind): Set before = SavedFields(id)
    If CurrentRevision(id) <> pdfRevision Then Err.Raise vbObjectError + 4722, , "Target draft changed during preview. Preview again."
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In incoming.Keys
        If (mode = "Replace mapped values" Or Len(Nz(before(key), "")) = 0) And Not SameFieldValue(before(key), incoming(key)) Then result.Add CStr(key), incoming(key)
    Next
    pdfTarget = id: pdfChoice = century & "|" & identifierKind & "|" & mode
    Set pdfChanges = result: Set PDFPrepareMerge = result
End Function
Public Function PDFExtractionStatus(Optional ByVal testPath As String = "") As String
    Dim p As Object, j As Object, x As Object, text As String, before As String
    On Error GoTo Failed
    Set p = CreateObject("AcroExch.PDDoc")
    text = "Acrobat automation available; extraction has not been tested with a PDF in this session."
    If Len(testPath) > 0 Then
        before = SourceFileHash(testPath)
        If Not p.Open(testPath) Then Err.Raise 5, , "Cannot open the test PDF."
        Set j = p.GetJSObject(): Set x = j.xfa
        If Len(CStr(x.template.saveXML())) = 0 Or Len(CStr(x.datasets.saveXML())) = 0 Then Err.Raise 5, , "No XFA template/data extraction."
        text = "PASS: read-only XFA extraction; Acrobat " & CStr(j.app.viewerVersion)
        p.Close
        If SourceFileHash(testPath) <> before Then Err.Raise 5, , "Test PDF changed."
    End If
    Set p = Nothing: PDFExtractionStatus = text: Exit Function
Failed:
    text = "PDF extraction unavailable: " & Err.Description & ". Requires desktop Adobe Acrobat with working AcroExch.PDDoc / GetJSObject / XFA extraction. Reader and browser viewers are insufficient."
    On Error Resume Next
    If Not p Is Nothing Then p.Close
    PDFExtractionStatus = text
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
