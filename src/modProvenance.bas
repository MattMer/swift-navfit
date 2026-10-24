Attribute VB_Name = "modProvenance"
Option Compare Database
Option Explicit
Public Function OpenImportProvenance(ByVal receiptID As Long) As Boolean
    Dim r As DAO.Recordset, context As String
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ImportReceipt WHERE ImportReceiptID=" & receiptID, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 1600, "modProvenance.OpenImportProvenance", "Select an existing import receipt first."
    context = "Receipt " & receiptID & " / Swift report " & r!ReportID & " / source report " & r!ExternalReportID & vbCrLf
    context = context & "Imported: " & r!ImportedAt & " / contributor " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & r!AuthorProfileID), "Unknown") & vbCrLf
    context = context & "Source: " & Nz(r!SourcePath, "") & vbCrLf & "SHA256: " & r!SourceHash & vbCrLf
    context = context & "Adapter: " & Nz(r!AdapterVersion, "") & " / encoding: " & Nz(r!ValueEncoding, "") & vbCrLf
    context = context & "Before revision: " & Nz(r!BeforeRevisionID, "None recorded") & " / applied revision: " & Nz(r!AppliedRevisionID, "None recorded") & " / undone: " & CStr(Nz(r!IsUndone, False))
    r.Close
    DoCmd.OpenForm "frmProvenance"
    Forms!frmProvenance!txtReceiptID = receiptID
    Forms!frmProvenance!txtContext = context
    Forms!frmProvenance!cboTable = "Reports"
    Forms!frmProvenance!txtRow = Null
    Forms!frmProvenance!txtSearch = ""
    Call LoadProvenanceFields()
    OpenImportProvenance = True
End Function
Public Function ProvenanceFromOperations() As Boolean
    On Error GoTo Failed
    ProvenanceFromOperations = OpenImportProvenance(Nz(Forms!frmOperations!lstReceipts, 0))
    Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Import provenance"
End Function
Public Function LoadProvenanceFields() As Boolean
    Dim sql As String, table As String, row As Variant, term As String, id As Long
    id = Nz(Forms!frmProvenance!txtReceiptID, 0)
    table = Nz(Forms!frmProvenance!cboTable, "All")
    row = Forms!frmProvenance!txtRow
    term = Left$(Nz(Forms!frmProvenance!txtSearch, ""), 255)
    sql = "SELECT ImportRawValueID,SourceTable,SourceRow,FieldOrdinal,FieldKey,DAOType,IsNullValue FROM ImportRawValue WHERE ImportReceiptID=" & id
    If table <> "All" Then
        If table <> "Reports" And table <> "Folders" And table <> "Summary" Then Err.Raise vbObjectError + 1601, "modProvenance.LoadProvenanceFields", "Choose a listed source table."
        sql = sql & " AND SourceTable='" & table & "'"
    End If
    If Not IsNull(row) And Nz(row, "") <> "" Then
        If Not IsNumeric(row) Then Err.Raise vbObjectError + 1602, "modProvenance.LoadProvenanceFields", "Source row must be a nonnegative whole number."
        If CDbl(row) < 0 Or CDbl(row) > 2147483647# Or CDbl(row) <> Fix(CDbl(row)) Then Err.Raise vbObjectError + 1602, "modProvenance.LoadProvenanceFields", "Source row must be a nonnegative whole number."
        sql = sql & " AND SourceRow=" & CLng(row)
    End If
    sql = sql & " AND InStr(1,FieldKey,'" & Replace(term, "'", "''") & "',1)>0 ORDER BY SourceTable,SourceRow,FieldOrdinal"
    Forms!frmProvenance!lstFields.RowSource = sql
    Forms!frmProvenance!lstFields.Requery
    Forms!frmProvenance!txtValue = Null
    Forms!frmProvenance!lblValue.Caption = "Select a retained field. Source row is a capture ordinal, not the NAVFIT report ID."
    Forms!frmProvenance!lblCount.Caption = Forms!frmProvenance!lstFields.ListCount & " retained fields shown"
    LoadProvenanceFields = True
End Function
Public Function FilterProvenanceUI() As Boolean
    On Error GoTo Failed
    FilterProvenanceUI = LoadProvenanceFields(): Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Import provenance"
End Function
Public Function ShowProvenanceValue() As Boolean
    Dim r As DAO.Recordset, id As Long, receipt As Long, state As String
    id = Nz(Forms!frmProvenance!lstFields, 0)
    receipt = Nz(Forms!frmProvenance!txtReceiptID, 0)
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ImportRawValue WHERE ImportRawValueID=" & id & " AND ImportReceiptID=" & receipt, dbOpenSnapshot)
    Forms!frmProvenance!txtValue = Null
    If r.EOF Then r.Close: Exit Function
    If r!IsNullValue Then
        state = "NULL"
    ElseIf Len(Nz(r!RawValue, "")) = 0 Then
        state = "Empty text"
    Else
        state = Len(r!RawValue) & " stored characters"
    End If
    Forms!frmProvenance!lblValue.Caption = r!SourceTable & " / row " & r!SourceRow & " / " & r!FieldKey & " / DAO type " & r!DAOType & " / " & state
    Forms!frmProvenance!txtValue = r!RawValue
    r.Close: ShowProvenanceValue = True
End Function
