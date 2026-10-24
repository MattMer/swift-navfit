Attribute VB_Name = "NAVFITAdapterHarness"
Option Compare Binary
Option Explicit
Private checks As Long
Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 5200, , message
End Sub
Public Function TestNAVFITAdapter(ByVal sourcePath As String, ByVal outputPath As String) As String
    Dim source As DAO.Database, target As DAO.Database, r As DAO.Recordset, t As DAO.Recordset, values As Object, row As Variant, parts As Variant, k As Variant, n As Long, id As Long, incoming As Object, saved As Object, key As Variant
    On Error GoTo Failed
    checks = 0
    For Each k In Array("EVAL", "CHIEFEVAL", "FITREP")
        Check NAVFITMapRows(CStr(k)).Count = IIf(k = "CHIEFEVAL", 96, 75), "Incomplete saved-field coverage"
        For n = 0 To 5
            Check NAVFITEncode(NAVFITDecode(n, "trait"), "trait") = n, "Trait code round trip"
            Check NAVFITEncode(NAVFITDecode(n, "promotion"), "promotion") = n, "Promotion code round trip"
        Next
    Next
    Check NAVFITDecode("26SEP30", "counseling", 2000) = "2026-09-30", "Counseling decode"
    Check NAVFITEncode("2026-09-30", "counseling") = "26SEP30", "Counseling encode"
    For Each k In Array("NOT REQ", "NOT PERF")
        Check NAVFITDecode(CStr(k), "counseling") = CStr(k), "Counseling status decode without century"
        Check NAVFITEncode(CStr(k), "counseling") = CStr(k), "Counseling status encode"
    Next
    Check NAVFITMergeDecision("base", "base", "incoming") = "Accept returned", "Incoming edit not recognized"
    Check NAVFITMergeDecision("base", "local", "base") = "Keep current", "Local edit lost"
    Check NAVFITMergeDecision("base", "local", "local") = "Keep current", "Same edit conflict"
    Check NAVFITMergeDecision("base", "local", "incoming") = "Conflict", "Conflict missed"
    Check NAVFITMergeDecision(Null, "", "incoming") = "Conflict", "Null/empty conflict missed"
    Check NAVFITMergeDecision(Null, Null, "") = "Accept returned", "Null/empty change missed"
    Set source = DBEngine.OpenDatabase(sourcePath, False, True)
    FileCopy sourcePath, outputPath
    Set target = DBEngine.OpenDatabase(outputPath)
    Set r = source.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null ORDER BY ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        Set values = NAVFITReadFields(r, 2000)
        Set t = target.OpenRecordset("SELECT * FROM Reports WHERE ReportID=" & r!ReportID, dbOpenDynaset)
        t.Edit: NAVFITWriteFields t, values: t.Update
        For Each row In NAVFITMapRows(values("FormKind"))
            parts = Split(row, "|")
            Check SameFieldValue(r.Fields(parts(1)).Value, t.Fields(parts(1)).Value), "Native value changed: " & values("FormKind") & " " & parts(1)
        Next
        t.Close: r.MoveNext
    Loop
    r.Close: source.Close: target.Close
    Call ConfigureWorkspace("SYNTHETIC NAVFIT ADAPTER", "Synthetic Tester")
    Call OpenReviewFile(sourcePath)
    For n = 2 To 4
        Set incoming = FullNAVFITImportValues(n, 2000)
        id = ImportReviewedReport(n, incoming("SourceHash"), 2000)
        Set saved = SavedFields(id)
        For Each key In incoming.Keys
            If key <> "SourceHash" Then Check SameFieldValue(incoming(key), saved(key)), "Imported field mismatch " & key
        Next
        Check DCount("*", "ImportReceipt", "ReportID=" & id & " AND AdapterVersion='navfit-full-v1/local-CStr'") = 1, "Full receipt missing"
    Next
    Call CloseReviewFile
    TestNAVFITAdapter = "PASS: " & checks & " NAVFIT adapter assertions": Exit Function
Failed:
    TestNAVFITAdapter = "FAIL: " & Err.Number & " " & Err.Description
End Function
