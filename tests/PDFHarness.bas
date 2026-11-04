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
        id = ImportPDFPreview("", "SSN")
        If DCount("*", "Report") <> before + 1 Then Err.Raise 5, , "Automatic 20xx import failed"
        count = count + 1
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
    If Not TestPDFReapply(directory) Then Err.Raise 5, , "PDF undo/reapply regression"
    count = count + 9
    count = count + 1: TestPDFImport = "PASS: " & count & " native PDF import assertions"
    Exit Function
Failed:
    TestPDFImport = "FAIL after " & count & " assertions: " & Err.Number & " / " & Err.Description
End Function
Private Function TestPDFReapply(ByVal directory As String) As Boolean
    Dim preview As Object, changes As Object, receipt As Long, first As Long, second As Long, revision As Long, rawCount As Long, rejected As Boolean, path As String
    path = directory & "\synthetic-eval.pdf"
    receipt = DMin("ImportReceiptID", "ImportReceipt", "AdapterVersion Like 'pdf-xfa-v*'")
    first = DLookup("ReportID", "ImportReceipt", "ImportReceiptID=" & receipt)
    rawCount = DCount("*", "ImportRawValue", "ImportReceiptID=" & receipt)
    'Schema-43 receipts have no attempt field. Migration leaves the new field null.
    CurrentDb.Execute "UPDATE ImportReceipt SET ReceiptAttempt=Null WHERE ImportReceiptID=" & receipt, dbFailOnError
    If Not UndoNarrativeImport(receipt) Then Exit Function
    If Not DLookup("IsUndone", "ImportReceipt", "ImportReceiptID=" & receipt) Or Not DLookup("IsArchived", "Report", "ReportID=" & first) Then Exit Function
    second = CommitDraft(0, "EVAL", "Correct PDF destination", "Synthetic PDF reapply", Null, Null, "", "Regression")
    revision = CurrentRevision(second)
    Set preview = PreviewStructuredPDF(path): Set changes = PDFPrepareMerge("20", "SSN", second, "Fill empty fields")
    If ImportPDFPreview("20", "SSN", second, "Fill empty fields") <> second Then Exit Function
    If DCount("*", "ImportReceipt", "SourceHash=" & QuoteSQL(SourceFileHash(path))) <> 2 Then Exit Function
    If DMax("ReceiptAttempt", "ImportReceipt", "SourceHash=" & QuoteSQL(SourceFileHash(path))) <> 2 Then Exit Function
    If DCount("*", "ImportRawValue", "ImportReceiptID=" & receipt) <> rawCount Or Not DLookup("IsUndone", "ImportReceipt", "ImportReceiptID=" & receipt) Then Exit Function
    Set preview = PreviewStructuredPDF(path)
    On Error Resume Next
    Call ImportPDFPreview("20", "SSN"): rejected = (Err.Number = vbObjectError + 4713): Err.Clear
    On Error GoTo 0
    If Not rejected Then Exit Function
    receipt = DMax("ImportReceiptID", "ImportReceipt", "ReportID=" & second)
    If Not UndoNarrativeImport(receipt) Then Exit Function
    If DLookup("IsArchived", "Report", "ReportID=" & second) Or Not IsNull(DLookup("ShipStation", "Report", "ReportID=" & second)) Then Exit Function
    Set preview = PreviewStructuredPDF(path)
    If ImportPDFPreview("20", "SSN") <= second Then Exit Function
    If DMax("ReceiptAttempt", "ImportReceipt", "SourceHash=" & QuoteSQL(SourceFileHash(path))) <> 3 Then Exit Function
    TestPDFReapply = True
End Function
