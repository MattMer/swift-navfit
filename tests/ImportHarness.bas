Attribute VB_Name = "ImportHarness"
Option Explicit
Public Function TryHash(ByVal path As String) As String
    On Error GoTo Failed
    TryHash = SourceFileHash(path): Exit Function
Failed:
    TryHash = "FAIL: " & Err.Number & " " & Err.Description
End Function
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 700, , message
End Sub
Private Function Rejected(ByVal id As Long) As String
    On Error GoTo Failed
    Call ImportReviewedNarrative(id)
    Rejected = "NOT REJECTED": Exit Function
Failed:
    Rejected = Err.Description
End Function
Public Function TestImports(ByVal path As String) As String
    Dim source As DAO.Database, receipts As DAO.Recordset, raw As DAO.Recordset, stored As DAO.Recordset, report As DAO.Recordset
    Dim id As Long, externalID As Long, table As Variant, row As Long, i As Long, value As Variant, expected As String
    Dim failedPath As String, before As Long
    On Error GoTo Failed
    Call ConfigureWorkspace("SYNTHETIC IMPORT", "Synthetic Import Tester")
    Call OpenReviewFile(path)
    Check Forms!frmImportReview!cmdImport.OnClick = "=UXApplyFullNAVFIT()", "Import button is not wired"
    Set source = DBEngine.OpenDatabase(path, False, True)
    Check InStr(Rejected(1), "quarantined") > 0, "Null type accepted"
    Check DCount("*", "Report") = 0, "Quarantine created draft"
    'Late provenance failure must roll back nested draft/revision writes too.
    CurrentDb.TableDefs("ImportRawValue").Fields("FieldKey").ValidationRule = "<> 'Comments'"
    Check InStr(Rejected(2), "NOT REJECTED") = 0, "Injected failure accepted"
    Check DCount("*", "Report") = 0 And DCount("*", "ReportRevision") = 0 And DCount("*", "ImportReceipt") = 0 And DCount("*", "ImportRawValue") = 0, "Import failure left partial rows"
    CurrentDb.TableDefs("ImportRawValue").Fields("FieldKey").ValidationRule = ""
    For externalID = 2 To 4
        id = ImportReviewedNarrative(externalID, "Synthetic reviewed display")
        Set raw = source.OpenRecordset("SELECT * FROM Reports WHERE ReportID=" & externalID, dbOpenSnapshot)
        Set report = CurrentDb.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
        Check IsNull(raw!Comments) = IsNull(report!OfficialComments), "Narrative null changed"
        If Not IsNull(raw!Comments) Then Check StrComp(raw!Comments, report!OfficialComments, vbBinaryCompare) = 0, "Narrative changed"
        Check IsNull(report!PeriodStart) And IsNull(report!PeriodEnd) And IsNull(report!FolderID), "Unverified fields inferred"
        Check report!MemberName = "Synthetic reviewed display", "Reviewed display lost"
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Initial revision missing"
        raw.Close: report.Close
        Check InStr(Rejected(externalID), "already been imported") > 0, "Duplicate accepted"
    Next
    Check DCount("*", "Report") = 3 And DCount("*", "ImportReceipt") = 3, "Wrong import counts"
    Set receipts = CurrentDb.OpenRecordset("ImportReceipt", dbOpenSnapshot)
    Do While Not receipts.EOF
        Check receipts!SourceHash = SourceFileHash(path), "Wrong provenance hash"
        For Each table In Array("Folders", "Reports", "Summary")
            Set raw = source.OpenRecordset("SELECT * FROM [" & table & "]", dbOpenSnapshot)
            row = 0
            Do While Not raw.EOF
                Set stored = CurrentDb.OpenRecordset("SELECT * FROM ImportRawValue WHERE ImportReceiptID=" & receipts!ImportReceiptID & " AND SourceTable='" & table & "' AND SourceRow=" & row & " ORDER BY FieldOrdinal", dbOpenSnapshot)
                For i = 0 To raw.Fields.Count - 1
                    Check Not stored.EOF, "Missing raw field"
                    Check stored!FieldKey = raw.Fields(i).Name And stored!DAOType = raw.Fields(i).Type, "Raw field/type mismatch"
                    value = raw.Fields(i).Value
                    Check stored!IsNullValue = IsNull(value), "Raw null mismatch"
                    If Not IsNull(value) Then
                        If VarType(value) = vbDate Then expected = CStr(CDbl(value)) Else expected = CStr(value)
                        Check StrComp(Nz(stored!RawValue, ""), expected, vbBinaryCompare) = 0, "Raw value changed"
                    End If
                    stored.MoveNext
                Next
                Check stored.EOF, "Extra raw field"
                stored.Close: row = row + 1: raw.MoveNext
            Loop
            raw.Close
        Next
        receipts.MoveNext
    Loop
    receipts.Close: source.Close: Call CloseReviewFile()
    Check DCount("*", "ImportRawValue") = 2472, "Raw provenance count wrong"
    TestImports = "PASS: three narrative imports; quarantine, duplicates, rollback and all 2472 provenance values verified"
    Exit Function
Failed:
    TestImports = "FAIL: " & Err.Number & " " & Err.Description
End Function
