Attribute VB_Name = "PDFHarness"
Option Compare Binary
Option Explicit
Public Function TestPDFImport(ByVal directory As String) As String
    Dim kind As Variant, preview As Object, id As Long, failure As Boolean, count As Long, before As Long, sourceHash As String, path As String
    On Error GoTo Failed
    Call ConfigureWorkspace("SYNTHETIC COMMAND", "Synthetic PDF tester")
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        path = directory & "\synthetic-" & LCase$(kind) & ".pdf": sourceHash = SourceFileHash(path)
        Set preview = PreviewStructuredPDF(path)
        If preview("FormKind") <> kind Then Err.Raise 5, , "Wrong PDF kind"
        count = count + 1: before = DCount("*", "Report")
        On Error Resume Next
        id = ImportPDFPreview("", "SSN"): failure = Err.Number <> 0: Err.Clear
        On Error GoTo Failed
        If Not failure Or DCount("*", "Report") <> before Then Err.Raise 5, , "Century rejection failed"
        count = count + 1
        id = ImportPDFPreview("20", "SSN")
        If MajorScalar("SELECT FormKind FROM Report WHERE ReportID=" & id) <> kind Then Err.Raise 5, , "Wrong committed kind"
        If MajorScalar("SELECT ShipStation FROM Report WHERE ReportID=" & id) <> "SYNTHETIC UNIT" Then Err.Raise 5, , "Wrong committed unit"
        If MajorScalar("SELECT PeriodStart FROM Report WHERE ReportID=" & id) <> #1/1/2026# Then Err.Raise 5, , "Wrong committed date"
        If InStr(MajorScalar("SELECT OfficialComments FROM Report WHERE ReportID=" & id), "O'Brien & example") = 0 Then Err.Raise 5, , "Lost punctuation"
        If MajorScalar("SELECT MemberSSN FROM Report WHERE ReportID=" & id) <> "000000000" Then Err.Raise 5, , "Wrong identifier"
        count = count + 5
        Set preview = PreviewStructuredPDF(path)
        On Error Resume Next
        id = ImportPDFPreview("20", "SSN"): failure = Err.Number <> 0: Err.Clear
        On Error GoTo Failed
        If Not failure Or DCount("*", "Report") <> before + 1 Then Err.Raise 5, , "Duplicate rejection failed"
        If SourceFileHash(path) <> sourceHash Then Err.Raise 5, , "Source PDF changed"
        count = count + 2
    Next
    If DCount("*", "ImportRawValue", "SourceTable='PDF-XFA'") <> 6 Then Err.Raise 5, , "Missing XML provenance"
    count = count + 1: TestPDFImport = "PASS: " & count & " native PDF import assertions"
    Exit Function
Failed:
    TestPDFImport = "FAIL after " & count & " assertions: " & Err.Number & " / " & Err.Description
End Function
