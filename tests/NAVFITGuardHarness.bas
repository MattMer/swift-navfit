Attribute VB_Name = "NAVFITGuardHarness"
Option Compare Binary
Option Explicit
Private checks As Long
Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    checks = checks + 1
    If Not condition Then Err.Raise vbObjectError + 5800, , message
End Sub
Private Function PreviewFailure(ByVal file As String, ByVal baseline As String) As String
    On Error GoTo Failed
    Call NAVFITPreviewReturned(file, baseline)
    PreviewFailure = "ACCEPTED": Exit Function
Failed:
    PreviewFailure = CStr(Err.Number) & " " & Err.Source & ": " & Err.Description
End Function
Private Function ApplyFailure() As String
    On Error GoTo Failed
    Call NAVFITApplyReturned
    ApplyFailure = "ACCEPTED": Exit Function
Failed:
    ApplyFailure = Err.Description
End Function
Private Function ExportFailure(ByVal id As Long, ByVal file As String, ByVal template As String) As String
    On Error GoTo Failed
    Call NAVFITExportReport(id, file, template)
    ExportFailure = "ACCEPTED": Exit Function
Failed:
    ExportFailure = Err.Description
End Function
Private Function EncodeFailure(ByVal value As Variant, ByVal conversion As String) As String
    On Error GoTo Failed
    Call NAVFITEncode(value, conversion)
    EncodeFailure = "ACCEPTED": Exit Function
Failed:
    EncodeFailure = Err.Description
End Function
Private Function DecodeFailure(ByVal value As Variant, ByVal conversion As String) As String
    On Error GoTo Failed
    Call NAVFITDecode(value, conversion, 2000)
    DecodeFailure = "ACCEPTED": Exit Function
Failed:
    DecodeFailure = Err.Description
End Function
Public Function TestNAVFITGuards(ByVal directory As String, ByVal template As String) As String
    Dim file As String, baseline As String, bad As String, db As DAO.Database, changes As Object, rows As Collection, receipt As Long, revision As Long, raw As Long, oldValue As Variant, field As DAO.Field, columns As String, selected As String, source As String, missing As String, values As Object, decoded As Object, record As DAO.Recordset, id As Long, key As Variant
    On Error GoTo Failed
    file = directory & "\guard.accdb": baseline = file & ".swift-baseline.accdb"
    Call NAVFITExportReport(3, file, template)
    revision = CurrentRevision(3): raw = DCount("*", "ImportRawValue")
    Check InStr(ExportFailure(3, file, template), "not overwritten") > 0, "Existing export overwritten"
    Check EncodeFailure(Null, "flag") <> "ACCEPTED", "Unknown checkbox exported"
    Check EncodeFailure("5.5", "trait") <> "ACCEPTED", "Fractional trait exported"
    Check EncodeFailure("Unknown", "promotion") <> "ACCEPTED", "Unknown promotion exported"
    Check EncodeFailure("2026-02-30", "date") <> "ACCEPTED", "Impossible date exported"
    Check EncodeFailure("2100-01-01", "counseling") <> "ACCEPTED", "Unresolvable counseling century exported"
    Check NAVFITDecode(#1/1/2026 12:00:00 PM#, "date") = "2026-01-01", "Native timestamp did not normalize to calendar date"
    Check EncodeFailure(#1/1/2026 12:00:00 PM#, "date") <> "ACCEPTED", "Swift date time silently discarded"
    Set rows = NAVFITPreviewReturned(file, baseline)
    Set db = DBEngine.OpenDatabase(file): db.Execute "UPDATE Reports SET Comments='Changed after preview' WHERE ReportType Is Not Null", dbFailOnError: db.Close
    Check InStr(ApplyFailure(), "changed after preview") > 0, "Changed return file applied"
    Check CurrentRevision(3) = revision And DCount("*", "ImportRawValue") = raw, "Rejected return wrote values"
    Set rows = NAVFITPreviewReturned(file, baseline)
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "OfficialComments", "Local edit after preview"
    Call CommitSavedFields(3, CurrentRevision(3), changes, "Guard test local edit")
    Check InStr(ApplyFailure(), "Swift report changed") > 0, "Stale Swift preview applied"
    bad = directory & "\tampered-baseline.accdb": FileCopy baseline, bad
    Set db = DBEngine.OpenDatabase(bad): db.Execute "UPDATE SwiftValue SET ValueText='tampered' WHERE FieldKey='OfficialComments'", dbFailOnError: db.Close
    Check InStr(PreviewFailure(file, bad), "Baseline is modified") > 0, "Modified baseline accepted"
    bad = directory & "\wrong-identity.accdb": FileCopy file, bad
    Set db = DBEngine.OpenDatabase(bad)
    For Each field In db.TableDefs("Reports").Fields
        If columns <> "" Then columns = columns & ",": selected = selected & ","
        columns = columns & "[" & field.Name & "]"
        If field.Name = "ReportID" Then selected = selected & "900" Else selected = selected & "[" & field.Name & "]"
    Next
    db.Execute "INSERT INTO Reports (" & columns & ") SELECT " & selected & " FROM Reports WHERE ReportType Is Not Null", dbFailOnError
    db.Execute "DELETE FROM Reports WHERE ReportType Is Not Null AND ReportID<>900", dbFailOnError: db.Close
    Check InStr(PreviewFailure(bad, baseline), "identity differs") > 0, "Changed native identity result: " & PreviewFailure(bad, baseline)
    CurrentDb.Execute "UPDATE Report SET WorkflowState='Approved' WHERE ReportID=3", dbFailOnError
    Check InStr(PreviewFailure(file, baseline), "local approval") > 0, "Locked report accepted return"
    CurrentDb.Execute "UPDATE Report SET WorkflowState='Draft' WHERE ReportID=3", dbFailOnError
    Set rows = NAVFITPreviewReturned(file, baseline)
    Check NAVFITReturnedDecision("OfficialComments") = "Conflict", "Three-way conflict missing"
    NAVFITResolveReturned "OfficialComments", "Accept returned"
    Check NAVFITReturnedDecision("OfficialComments") = "Accept returned", "Resolved decision not exposed"
    oldValue = SavedFields(3)("OfficialComments")
    Check NAVFITApplyReturned() = 3, "Resolved return failed"
    receipt = DMax("ImportReceiptID", "ImportReceipt", "ReportID=3")
    Check UndoNarrativeImport(receipt), "Full returned import undo failed"
    Check SameFieldValue(SavedFields(3)("OfficialComments"), oldValue), "Undo lost prior local field"
    Check DLookup("IsUndone", "ImportReceipt", "ImportReceiptID=" & receipt), "Undo receipt not retained"
    source = DLookup("SourcePath", "ImportReceipt", "ReportID=2 AND IsUndone=False")
    missing = source & ".temporarily-unavailable"
    Name source As missing
    Check NAVFITExportReport(2, directory & "\source-independent.accdb", template) <> "", "Export depends on original source file"
    Name missing As source
    Set values = SavedFields(2)
    id = NAVFITCommitFields(0, values, "Synthetic new draft export", "New draft export guard")
    file = directory & "\new-draft.accdb"
    Check NAVFITExportReport(id, file, template) = file, "New draft export failed"
    Set db = DBEngine.OpenDatabase(file, False, True)
    Set record = db.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
    Set decoded = NAVFITReadFields(record, 2000)
    For Each key In decoded.Keys
        Check SameFieldValue(decoded(key), values(key)), "New draft export mismatch: " & key
    Next
    record.Close: db.Close
    Set rows = NAVFITPreviewReturned(file, file & ".swift-baseline.accdb")
    Check NAVFITApplyReturned() = id, "New draft baseline return failed"
    TestNAVFITGuards = "PASS: " & checks & " NAVFIT stale/hash/identity/lock/undo guards": Exit Function
Failed:
    TestNAVFITGuards = "FAIL: " & Err.Number & " " & Err.Description
End Function
