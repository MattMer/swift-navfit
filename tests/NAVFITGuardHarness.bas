Attribute VB_Name = "NAVFITGuardHarness"
Option Compare Binary
Option Explicit
Private checks As Long
Private lastExportFailureNumber As Long
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
    lastExportFailureNumber = Err.Number
    ExportFailure = Err.Description
End Function
Private Function GroupExportFailure(ByVal group As Long, ByVal file As String, ByVal template As String) As String
    Dim preview As Object
    On Error GoTo Failed
    Set preview = NAVFITGroupPreview(group, False)
    Call NAVFITExportGroup(group, file, template, preview("Token"), False)
    GroupExportFailure = "ACCEPTED": Exit Function
Failed:
    lastExportFailureNumber = Err.Number
    GroupExportFailure = Err.Description
End Function
Private Sub CheckExportDiagnostics(ByVal id As Long, ByVal directory As String, ByVal template As String)
    Dim db As DAO.Database, mode As Variant, destination As String, missing As String, failure As String, group As Long, preview As Object, revisions As Long, publishedHash As String
    Set db = CurrentDb: revisions = DCount("*", "ReportRevision")
    group = CreateSummaryGroup("SYNTHETIC DIAGNOSTIC GROUP", SavedFields(id)("FormKind"), "E7")
    AddGroupMember group, id, "E7"
    missing = directory & "\missing-123456789.accdb"
    For Each mode In Array("Report", "Group")
        destination = directory & "\diagnostic-" & mode & ".accdb"
        db.TableDefs("OperationJournal").Fields("OperationState").ValidationRule = "<> 'Failed'"
        If mode = "Report" Then failure = ExportFailure(id, destination, missing) Else failure = GroupExportFailure(group, destination, missing)
        db.TableDefs("OperationJournal").Fields("OperationState").ValidationRule = ""
        Check failure <> "ACCEPTED" And lastExportFailureNumber = 53, mode & " cleanup replaced the primary missing-template error: " & lastExportFailureNumber & " " & failure
        Check InStr(SwiftLastFailure(), "modNAVFITExchange.NAVFITExport" & mode & ":53") > 0, mode & " lost its primary diagnostic identity"
        Check InStr(SwiftLastFailure(), "record failed operation") > 0, mode & " omitted the separate journal cleanup failure"
        Check InStr(failure & SwiftLastFailure(), "123456789") = 0, mode & " diagnostics exposed an identifier"
        Check Len(Dir$(destination)) = 0 And Len(Dir$(destination & ".swift-baseline.accdb")) = 0, mode & " failure published output"
        Check DCount("*", "ReportRevision") = revisions, mode & " failed export changed saved revisions"
        If mode = "Report" Then
            Check NAVFITExportReport(id, destination, template) = destination, "Report retry failed"
        Else
            Set preview = NAVFITGroupPreview(group, False)
            Check NAVFITExportGroup(group, destination, template, preview("Token"), False) = destination, "Group retry failed"
        End If
        destination = directory & "\published-diagnostic-" & mode & ".accdb"
        db.TableDefs("OperationJournal").Fields("OperationState").ValidationRule = "<> 'Completed'"
        If mode = "Report" Then failure = ExportFailure(id, destination, template) Else failure = GroupExportFailure(group, destination, template)
        db.TableDefs("OperationJournal").Fields("OperationState").ValidationRule = ""
        Check failure <> "ACCEPTED" And lastExportFailureNumber = 3317, mode & " completion-journal failure changed its native error"
        Check InStr(SwiftLastFailure(), "files published") > 0, mode & " diagnostic falsely claimed output was not published"
        Check Len(Dir$(destination)) > 0 And Len(Dir$(destination & ".swift-baseline.accdb")) > 0, mode & " journal failure lost published files"
        publishedHash = SourceFileHash(destination) & SourceFileHash(destination & ".swift-baseline.accdb")
        If mode = "Report" Then failure = ExportFailure(id, destination, template) Else failure = GroupExportFailure(group, destination, template)
        Check InStr(failure, "not overwritten") > 0, mode & " retry replaced published files"
        Check publishedHash = SourceFileHash(destination) & SourceFileHash(destination & ".swift-baseline.accdb"), mode & " retry changed published bytes"
        Check DCount("*", "ReportRevision") = revisions, mode & " completion-journal failure changed revisions"
    Next
End Sub
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
    Dim chiefID As Long, nativeComments As Object, row As Variant, parts As Variant
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
    source = ApplyFailure()
    Check InStr(source, "Source changed.") > 0, "Changed return file guard: " & source
    Check CurrentRevision(3) = revision And DCount("*", "ImportRawValue") = raw, "Rejected return wrote values"
    Set rows = NAVFITPreviewReturned(file, baseline)
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "OfficialComments", "Local edit after preview"
    Call CommitSavedFields(3, CurrentRevision(3), changes, "Guard test local edit")
    revision = CurrentRevision(3): source = ApplyFailure()
    Check InStr(source, "Target revision changed.") > 0, "Stale Swift preview guard: " & source
    Check CurrentRevision(3) = revision And SavedFields(3)("OfficialComments") = "Local edit after preview", "Rejected stale return changed report"
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
    Set values = SavedFields(id): values("FormKind") = "CHIEFEVAL"
    Set nativeComments = CreateObject("Scripting.Dictionary")
    For Each row In NAVFITMapRows("CHIEFEVAL")
        parts = Split(row, "|")
        If parts(0) Like "Trait##Comment#" Then
            values(parts(0)) = Null: nativeComments.Add parts(0), parts(1)
        End If
    Next
    chiefID = NAVFITCommitFields(0, values, "Synthetic Chief member", "New Chief comment guard")
    revision = CurrentRevision(chiefID)
    Check DCount("*", "RevisionChange", "RevisionID=" & revision & " AND FieldKey Like 'Trait##Comment#'") = 0, "New Chief logged untouched null comments"
    bad = directory & "\chief-comment-guard.accdb": FileCopy file, bad
    Set db = DBEngine.OpenDatabase(bad)
    Set record = db.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenDynaset)
    record.Edit
    For Each key In nativeComments.Keys
        record.Fields(nativeComments(key)).Value = "Synthetic native " & key
    Next
    record.Update
    Set values = SavedFields(chiefID)
    record.Edit: NAVFITWriteFields record, values, revision: record.Update
    For Each key In nativeComments.Keys
        Check SameFieldValue(record.Fields(nativeComments(key)).Value, "Synthetic native " & key), "Untouched Chief comment erased: " & key
    Next
    Set changes = CreateObject("Scripting.Dictionary")
    changes.Add "Trait33Comment1", "Synthetic authored Chief comment"
    changes.Add "Trait33Comment2", "Synthetic comment before clear"
    Call CommitSavedFields(chiefID, revision, changes, "Author Chief comments")
    Set changes = CreateObject("Scripting.Dictionary")
    changes.Add "Trait33Comment2", Null: changes.Add "Trait33Comment3", ""
    Call CommitSavedFields(chiefID, CurrentRevision(chiefID), changes, "Clear Chief comments")
    Set values = SavedFields(chiefID)
    record.Edit: NAVFITWriteFields record, values, CurrentRevision(chiefID): record.Update
    Check SameFieldValue(record.Fields(nativeComments("Trait33Comment1")).Value, "Synthetic authored Chief comment"), "Authored Chief comment did not overwrite native text"
    Check IsNull(record.Fields(nativeComments("Trait33Comment2")).Value), "Logged Chief clear did not overwrite native text"
    Check SameFieldValue(record.Fields(nativeComments("Trait33Comment3")).Value, ""), "Explicit empty Chief comment did not overwrite native text"
    For Each key In nativeComments.Keys
        If key <> "Trait33Comment1" And key <> "Trait33Comment2" And key <> "Trait33Comment3" Then
            Check SameFieldValue(record.Fields(nativeComments(key)).Value, "Synthetic native " & key), "Chief save erased an untouched native comment: " & key
        End If
    Next
    record.Close: db.Close
    CheckExportDiagnostics id, directory, template
    TestNAVFITGuards = "PASS: " & checks & " NAVFIT stale/hash/identity/lock/undo and export diagnostic guards": Exit Function
Failed:
    TestNAVFITGuards = "FAIL: " & Err.Number & " " & Err.Description
End Function
