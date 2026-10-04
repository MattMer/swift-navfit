Attribute VB_Name = "modPDFImport"
Option Compare Binary
Option Explicit
Private pdfPreview As Object
Private pdfXML As String
Private pdfPath As String
Private pdfHash As String
Public Function PreviewStructuredPDF(ByVal path As String) As Object
    Dim p As Object, j As Object, x As Object, dom As Object, nodes As Object, result As Object, fields As Variant, key As Variant, templateHash As String, kind As String, value As Variant
    Dim failure As Long, message As String
    On Error GoTo Failed
    Set pdfPreview = Nothing: pdfXML = "": pdfPath = "": pdfHash = ""
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
    If kind = "CHIEFEVAL" Then
        fields = Array("NAMEmaster", "RATINGmaster", "DESIGNATIONmaster", "SSNDODIDmaster", "SHIPSTATION", "FromDate", "ToDate", "REPORTINGSENIORCOMMENTS")
    Else
        fields = Array("a", "b", "c", "d", "g", "k", "l", "H")
    End If
    For Each key In fields
        Set nodes = dom.selectNodes("//*[local-name()='" & key & "']")
        If nodes.length > 1 Then Err.Raise vbObjectError + 4705, , "Repeated PDF fields require manual review."
        If nodes.length = 0 Then value = Null Else value = CStr(nodes.Item(0).Text)
        result.Add CStr(key), value
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
Public Function ImportPDFPreview(ByVal century As String, ByVal identifierKind As String) As Long
    Dim fields As Variant, id As Long, changes As Object, receipt As DAO.Recordset, raw As DAO.Recordset, receiptID As Long, w As DAO.Workspace, failure As Long, message As String, identifier As Variant, db As DAO.Database, stage As String, expected As Long
    If pdfPreview Is Nothing Then Err.Raise vbObjectError + 4712, , "Preview a supported PDF first."
    If SourceFileHash(pdfPath) <> pdfHash Then Err.Raise vbObjectError + 4706, , "PDF changed after preview."
    If DCount("*", "ImportReceipt", "SourceHash=" & QuoteSQL(pdfHash) & " AND AdapterVersion='pdf-xfa-v1'") > 0 Then Err.Raise vbObjectError + 4713, , "This exact PDF has already been imported."
    If identifierKind <> "SSN" And identifierKind <> "DoDID" And identifierKind <> "Retain only" Then Err.Raise vbObjectError + 4714, , "Choose how to interpret the combined identifier field."
    If pdfPreview("FormKind") = "CHIEFEVAL" Then
        fields = Array("NAMEmaster", "RATINGmaster", "DESIGNATIONmaster", "SSNDODIDmaster", "SHIPSTATION", "FromDate", "ToDate", "REPORTINGSENIORCOMMENTS")
    Else
        fields = Array("a", "b", "c", "d", "g", "k", "l", "H")
    End If
    Set changes = CreateObject("Scripting.Dictionary")
    changes.Add "OfficialName", pdfPreview(fields(0)): changes.Add "RankRate", pdfPreview(fields(1)): changes.Add "Designator", pdfPreview(fields(2))
    changes.Add "ShipStation", pdfPreview(fields(4)): changes.Add "PeriodStart", PDFDate(pdfPreview(fields(5)), century): changes.Add "PeriodEnd", PDFDate(pdfPreview(fields(6)), century)
    changes.Add "OfficialComments", pdfPreview(fields(7)): identifier = pdfPreview(fields(3))
    If identifierKind = "SSN" Then changes.Add "MemberSSN", identifier
    If identifierKind = "DoDID" Then changes.Add "MemberDoDID", identifier
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    stage = "Create draft"
    id = CommitDraft(0, pdfPreview("FormKind"), "Imported PDF draft", Left$(Nz(pdfPreview(fields(0)), ""), 120), Null, Null, "Other PDF fields are retained as source XML and require review.", "PDF import")
    stage = "Read created revision": expected = CurrentRevision(id)
    stage = "Commit PDF fields"
    Call CommitSavedFields(id, expected, changes, "Reviewed structured PDF import")
    stage = "Retain receipt": Set db = CurrentDb
    Set receipt = db.OpenRecordset("ImportReceipt", dbOpenDynaset)
    receipt.AddNew: receipt!ReceiptGuid = NewIdentity(): receipt!SourceHash = pdfHash: receipt!SourcePath = pdfPath: receipt!ExternalReportID = 0: receipt!ReportID = id
    receipt!AuthorProfileID = MajorAuthor(): receipt!ImportedAt = Now(): receipt!AdapterVersion = "pdf-xfa-v1": receipt!ValueEncoding = "XFA XML; explicit century and identifier": receipt!AppliedRevisionID = CurrentRevision(id): receipt!IsUndone = False
    receipt.Update: receipt.Bookmark = receipt.LastModified: receiptID = receipt!ImportReceiptID: receipt.Close
    stage = "Retain XML"
    Set raw = db.OpenRecordset("ImportRawValue", dbOpenDynaset)
    raw.AddNew: raw!ImportReceiptID = receiptID: raw!SourceTable = "PDF-XFA": raw!SourceRow = 0: raw!FieldOrdinal = 0: raw!FieldKey = "datasets": raw!DAOType = dbMemo: raw!RawValue = pdfXML: raw!IsNullValue = False: raw.Update
    raw.AddNew: raw!ImportReceiptID = receiptID: raw!SourceTable = "PDF-XFA": raw!SourceRow = 0: raw!FieldOrdinal = 1: raw!FieldKey = "template-hash": raw!DAOType = dbText: raw!RawValue = pdfPreview("TemplateHash"): raw!IsNullValue = False: raw.Update: raw.Close
    If SourceFileHash(pdfPath) <> pdfHash Then Err.Raise vbObjectError + 4706, , "PDF changed during import."
    w.CommitTrans: ImportPDFPreview = id: Set pdfPreview = Nothing: Exit Function
Failed:
    failure = Err.Number: message = stage & ": " & Err.Description
    On Error Resume Next
    w.Rollback
    If Not receipt Is Nothing Then receipt.Close
    If Not raw Is Nothing Then raw.Close
    On Error GoTo 0
    Err.Raise failure, "ImportPDFPreview", message
End Function
Public Function ChoosePDFPreview() As Boolean
    Dim dialog As Object, result As Object, key As Variant, text As String
    On Error GoTo Failed
    Set dialog = Application.FileDialog(3): dialog.Title = "Preview a completed NAVPERS PDF": dialog.AllowMultiSelect = False
    dialog.Filters.Clear: dialog.Filters.Add "PDF forms", "*.pdf"
    If dialog.Show <> -1 Then Exit Function
    Set result = PreviewStructuredPDF(dialog.SelectedItems(1))
    For Each key In result.Keys: text = text & key & ": " & Nz(result(key), "[missing]") & vbCrLf: Next
    Forms!frmPDFImport!txtPreview = text: ChoosePDFPreview = True: Exit Function
Failed:
    Forms!frmPDFImport!lblStatus.Caption = Err.Description
End Function
Public Function ImportPDFUI() As Boolean
    Dim id As Long
    On Error GoTo Failed
    id = ImportPDFPreview(Nz(Forms!frmPDFImport!cboCentury, ""), Nz(Forms!frmPDFImport!cboIdentifier, ""))
    Forms!frmPDFImport!lblStatus.Caption = "Created draft " & id & ". Remaining fields need review; source XML retained."
    ImportPDFUI = True: Exit Function
Failed:
    Forms!frmPDFImport!lblStatus.Caption = Err.Description
End Function
