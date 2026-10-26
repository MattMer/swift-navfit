Attribute VB_Name = "MalformedImportHarness"
Option Explicit
' Malformed-input checks for the file importers (threat model recommendation 2).
' Every malformed case must be rejected with a message, quickly, without writing
' rows, changing the source file, leaving importer state or leaving temp files.
' Control cases must load. Fixtures come from scripts/test-malformed-imports.ps1
' and, for NAVFIT98A databases, from Swift's own embedded blank template.
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Const CASE_LIMIT_MS As Long = 20000
Private started As Long, beforeCounts As String, beforeHash As String, beforeTemp As Long

Private Function Counts() As String
    Counts = DCount("*", "Report") & "|" & DCount("*", "ReportRevision") & "|" & DCount("*", "CommandMember") & "|" & DCount("*", "ImportReceipt")
End Function
Private Function TempLeftovers() As Long
    Dim f As String
    f = Dir$(Environ$("TEMP") & "\SwiftNAVFIT-*")
    Do While Len(f) > 0: TempLeftovers = TempLeftovers + 1: f = Dir$(): Loop
End Function
Private Function FileHash(ByVal path As String) As String
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then FileHash = SourceFileHash(path)
End Function
Private Sub Begin(ByVal path As String)
    started = GetTickCount(): beforeCounts = Counts(): beforeHash = FileHash(path): beforeTemp = TempLeftovers()
End Sub
' expect "" = must load; "*" = any rejection; other text = the rejection must contain it.
Private Function Verdict(ByVal path As String, ByVal expect As String, ByVal loaded As Boolean, ByVal message As String, ByVal stateClear As Boolean) As String
    Dim elapsed As Long
    elapsed = GetTickCount() - started
    If elapsed > CASE_LIMIT_MS Then Verdict = "FAIL: took " & elapsed & " ms": Exit Function
    If Len(expect) = 0 And Not loaded Then Verdict = "FAIL: control rejected: " & message: Exit Function
    If Len(expect) > 0 And loaded Then Verdict = "FAIL: malformed input accepted": Exit Function
    If Len(expect) > 0 And Len(Trim$(message)) = 0 Then Verdict = "FAIL: rejected without a message": Exit Function
    If Len(expect) > 0 And expect <> "*" And InStr(1, message, expect, vbTextCompare) = 0 Then Verdict = "FAIL: unexpected rejection: " & message: Exit Function
    If Counts() <> beforeCounts Then Verdict = "FAIL: rows written (" & beforeCounts & " -> " & Counts() & ")": Exit Function
    If FileHash(path) <> beforeHash Then Verdict = "FAIL: source file changed": Exit Function
    If TempLeftovers() > beforeTemp Then Verdict = "FAIL: temporary file left behind": Exit Function
    If Not stateClear Then Verdict = "FAIL: importer state not cleared": Exit Function
    Verdict = "PASS: " & IIf(loaded, "loaded", "rejected - " & message) & " (" & elapsed & " ms)"
End Function

Public Function RUADKeysLine() As String
    RUADKeysLine = Join(RUADKeys(), ",")
End Function

' PDF: the local reader, then the full preview, which must also clear its context.
Public Function MalformedPDF(ByVal path As String, ByVal expect As String) As String
    Dim packets As Object, loaded As Boolean, message As String, preview As Object
    Begin path
    On Error Resume Next
    Set packets = PDFLocalPackets(path)
    If Err.Number = 0 Then
        loaded = packets.Exists("template") And packets.Exists("datasets")
        If Not loaded Then message = "packets missing"
    Else
        message = Err.Description
    End If
    Err.Clear
    If Not loaded Then Set preview = PreviewStructuredPDF(path): Err.Clear
    On Error GoTo 0
    MalformedPDF = Verdict(path, expect, loaded, message, Len(PDFImportPath()) = 0)
End Function

' RUAD CSV: load only (Apply is never called). Optionally confirm a value is kept literally.
Public Function MalformedCSV(ByVal path As String, ByVal expect As String, Optional ByVal field As String = "", Optional ByVal literal As String = "") As String
    Dim loaded As Boolean, message As String, value As String
    Begin path
    On Error Resume Next
    loaded = (RUADLoad(path) > 0)
    If Err.Number <> 0 Then message = Err.Description: loaded = False
    If loaded And Len(field) > 0 Then value = RUADRow(1)("Data")(field)
    Err.Clear
    On Error GoTo 0
    If loaded And Len(field) > 0 And value <> literal Then MalformedCSV = "FAIL: " & field & " not kept literally: " & value: Call RUADReset: Exit Function
    MalformedCSV = Verdict(path, expect, loaded, message, True)
    Call RUADReset
    If RUADCount() <> 0 Then MalformedCSV = "FAIL: RUAD rows kept after reset"
End Function

' NAVFIT98A review import (OpenReviewFile) or the native exchange check (ValidateNAVFITDatabase).
Public Function MalformedNAVFIT(ByVal path As String, ByVal expect As String, ByVal exchange As Boolean) As String
    Dim loaded As Boolean, message As String, db As DAO.Database
    Begin path
    On Error Resume Next
    If exchange Then
        Set db = DBEngine.OpenDatabase(path, False, True)
        If Err.Number = 0 Then ValidateNAVFITDatabase db
    Else
        Call OpenReviewFile(path, False)
    End If
    If Err.Number = 0 Then loaded = True Else message = Err.Description
    Err.Clear
    If Not db Is Nothing Then db.Close: Set db = Nothing
    If Not exchange Then Call CloseReviewFile
    Err.Clear
    On Error GoTo 0
    MalformedNAVFIT = Verdict(path, expect, loaded, message, Len(ReviewPreparedCopyPath()) = 0)
End Function

' Write Swift's embedded NAVFIT98A blank to folder\nav-control.accdb and derive broken copies.
Public Function MakeNAVFITCases(ByVal folder As String) As String
    Dim valid As String, db As DAO.Database, dom As Object, node As Object, stream As Object
    valid = folder & "\nav-control.accdb"
    Set dom = CreateObject("MSXML2.DOMDocument.6.0"): Set node = dom.createElement("template")
    node.DataType = "bin.base64": node.Text = NAVFITTemplateData()
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 1: stream.Open
    stream.Write node.nodeTypedValue: stream.SaveToFile valid, 2: stream.Close
    If SourceFileHash(valid) <> NAVFITTemplateHash() Then Err.Raise 5, "MakeNAVFITCases", "Embedded NAVFIT template did not verify."
    Mutate valid, folder & "\nav-missing-table.accdb", "DROP TABLE Reports"
    Mutate valid, folder & "\nav-missing-field.accdb", "ALTER TABLE Reports DROP COLUMN FullName"
    Mutate valid, folder & "\nav-wrong-type.accdb", "ALTER TABLE Reports DROP COLUMN FullName", "ALTER TABLE Reports ADD COLUMN FullName LONG"
    Mutate valid, folder & "\nav-extra-field.accdb", "ALTER TABLE Reports ADD COLUMN Injected TEXT(10)"
    Mutate valid, folder & "\exch-extra-table.accdb", "CREATE TABLE Injected (Id LONG)"
    FileCopy valid, folder & "\exch-saved-query.accdb"
    Set db = DBEngine.OpenDatabase(folder & "\exch-saved-query.accdb")
    db.CreateQueryDef "Injected", "SELECT * FROM Reports": db.Close
    FileCopy valid, folder & "\exch-control.accdb"
    DBEngine.CompactDatabase valid, folder & "\nav-password.accdb", dbLangGeneral & ";pwd=Synthetic1"
    MakeNAVFITCases = "PASS: NAVFIT98A fixtures written"
End Function
Private Sub Mutate(ByVal source As String, ByVal target As String, ParamArray statements() As Variant)
    Dim db As DAO.Database, s As Variant
    FileCopy source, target
    Set db = DBEngine.OpenDatabase(target)
    For Each s In statements: db.Execute CStr(s), dbFailOnError: Next
    db.Close
End Sub
