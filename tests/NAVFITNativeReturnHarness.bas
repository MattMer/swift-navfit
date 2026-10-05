Attribute VB_Name = "NAVFITNativeReturnHarness"
Option Compare Binary
Option Explicit
Public Function VerifyNativeReturn(ByVal reportID As Long, ByVal path As String, ByVal template As String, ByVal marker As String, Optional ByVal alreadyApplied As Boolean = False) As String
    Dim rows As Collection, row As Variant, incoming As Object, native As DAO.Database, r As DAO.Recordset, key As Variant, saved As Object, reexport As String, a As DAO.Recordset, b As DAO.Recordset, output As DAO.Database, checks As Long
    On Error GoTo Failed
    Set native = DBEngine.OpenDatabase(path, False, True)
    Set r = native.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
    Set incoming = NAVFITReadFields(r, 2000)
    If marker <> "" And InStr(incoming("OfficialComments"), marker) = 0 Then Err.Raise vbObjectError + 5700, , "Native UI edit is absent."
    r.Close: native.Close
    If Not alreadyApplied Then
        Set rows = NAVFITPreviewReturned(path, path & ".swift-baseline.accdb")
        For Each row In rows
            If Right$(row, 9) = "|Conflict" Then Err.Raise vbObjectError + 5700, , "Unexpected local conflict: " & row
        Next
        If NAVFITApplyReturned() <> reportID Then Err.Raise vbObjectError + 5700, , "Wrong return target."
    End If
    Set saved = SavedFields(reportID)
    For Each key In incoming.Keys
        checks = checks + 1
        If Not SameFieldValue(incoming(key), saved(key)) Then Err.Raise vbObjectError + 5700, , "Returned field mismatch: " & key
    Next
    reexport = path & "." & Replace(Replace(NewIdentity(), "{", ""), "}", "") & ".reexport.accdb"
    Call NAVFITExportReport(reportID, reexport, template)
    Set native = DBEngine.OpenDatabase(path, False, True): Set output = DBEngine.OpenDatabase(reexport, False, True)
    Set a = native.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
    Set b = output.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
    For Each key In a.Fields
        checks = checks + 1
        If Not SameFieldValue(key.Value, b.Fields(key.Name).Value) Then Err.Raise vbObjectError + 5700, , "Re-export lost native value: " & key.Name
    Next
    a.Close: b.Close: native.Close: output.Close
    VerifyNativeReturn = "PASS: " & checks & " actual native return/re-export comparisons for " & incoming("FormKind") & "; reexport=" & reexport: Exit Function
Failed:
    VerifyNativeReturn = "FAIL: " & Err.Number & " " & Err.Description
End Function
