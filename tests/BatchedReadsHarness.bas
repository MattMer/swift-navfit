Attribute VB_Name = "BatchedReadsHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal value As Boolean, ByVal message As String)
    If Not value Then Err.Raise vbObjectError + 9300, , message
    checks = checks + 1
End Sub
Private Function ScalarStamp(ByVal selection As String) As String
    Dim part As Variant, selected As New Collection, id As Long, r As DAO.Recordset
    If Len(selection) = 0 Then Err.Raise vbObjectError + 1400, , "Select reports first."
    For Each part In Split(selection, ",")
        id = CLng(part)
        If id <= 0 Or CStr(id) <> CStr(part) Then Err.Raise vbObjectError + 1401, , "Invalid report selection."
        selected.Add id, CStr(id)
    Next
    For Each part In selected
        Set r = CurrentDb.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & part, dbOpenSnapshot)
        If r.EOF Then Err.Raise vbObjectError + 1402, , "Selected report missing."
        If r!IsArchived Or r!WorkflowState <> "Draft" Then Err.Raise vbObjectError + 1403, , "Selected report is archived or locked."
        ScalarStamp = ScalarStamp & part & ":" & Nz(r!FolderID, 0) & ":" & Nz(r!CurrentRevisionID, 0) & ":" & Nz(DMax("FolderMoveID", "FolderMove", "ReportID=" & part), 0) & ";": r.Close
    Next
End Function
Private Function StampError(ByVal selection As String, ByVal batch As Boolean) As Long
    Dim value As String
    On Error GoTo Failed
    If batch Then value = FolderBatchStamp(selection) Else value = ScalarStamp(selection)
    Exit Function
Failed:
    StampError = Err.Number
End Function
Private Function CacheError(ByVal id As Long, ByVal batch As Boolean) As Long
    Dim rows As DAO.Recordset
    On Error GoTo Failed
    If batch Then
        Set rows = CurrentDb.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
        AssertReportCacheRows rows: rows.Close
    Else
        AssertReportCache id
    End If
    Exit Function
Failed:
    CacheError = Err.Number
End Function
Private Sub CompareCache(ByVal id As Long, ByVal expected As Long)
    Check CacheError(id, False) = expected, "Scalar cache rejection changed"
    Check CacheError(id, True) = expected, "Batch cache rejection differs"
End Sub
Private Function PathError(ByVal id As Long, ByVal batch As Boolean) As Long
    Dim value As String, snapshot As Object
    On Error GoTo Failed
    If batch Then
        Set snapshot = FolderPathSnapshot(): value = FolderPathFromSnapshot(id, snapshot)
    Else
        value = FolderPath(id)
    End If
    Exit Function
Failed:
    PathError = Err.Number
End Function
Public Function TestBatchedReads() As String
    Dim r As DAO.Recordset, rows As Object, folders As Object, id As Long, rev As Long, version As Long, traits As Object, key As Variant
    Dim value As Variant, original As String, operation As Long, before As String, count As Long, failure As Long, parent As Variant, target As Long, selection As String, recovery As Long, keys As Variant, controls As Variant, i As Long
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("RecoveryDraft", dbOpenDynaset)
    r.AddNew: r!SessionGuid = NewIdentity(): r!AuthorProfileID = ActiveContributorID(): r!ReportID = 1: r!BaseRevisionID = CurrentRevision(1): r!CheckpointAt = #1/1/2000#: r!DraftTitle = "SYNTHETIC UNSAVED"
    keys = DraftFieldKeys(): controls = CatalogueRecoveryKeys()
    For i = 0 To UBound(keys)
        If keys(i) = "MemberSSN" Then r.Fields(CStr(controls(i))) = "SYNTHETIC INVALID RAW"
    Next
    r.Update: r.Bookmark = r.LastModified: recovery = r!RecoveryDraftID: r.Close
    operation = BeginOperation("SYNTHETIC interrupted", "Read-only preservation fixture")
    before = ReadHistoryCounts()
    Set r = RSGroupRows(1): r.Close
    Set rows = PriorEvaluationITARows(1): Set rows = UXHistoryITARows(1)
    original = FolderBatchStamp("1,2,3")
    Check ReadHistoryCounts() = before, "Read snapshots wrote data/history"
    Check DLookup("DraftTitle", "RecoveryDraft", "RecoveryDraftID=" & recovery) = "SYNTHETIC UNSAVED" And DLookup("CheckpointAt", "RecoveryDraft", "RecoveryDraftID=" & recovery) = #1/1/2000#, "Read snapshots changed recovery values/time"
    Check DLookup("OperationState", "OperationJournal", "OperationID=" & operation) = "Started", "Read snapshots rewrote interrupted history"
    Check FolderBatchStamp("3,2,1") = ScalarStamp("3,2,1"), "Stamp order/null folder/move identity differs"
    For Each value In Array("", "1,1", "01", "0", "999999", "1, 2")
        Check StampError(CStr(value), True) = StampError(CStr(value), False), "Selection rejection differs"
    Next
    Set folders = FolderPathSnapshot()
    Set r = CurrentDb.OpenRecordset("SELECT FolderID FROM Folder", dbOpenSnapshot)
    Do While Not r.EOF
        Check FolderPathFromSnapshot(r!FolderID, folders) = FolderPath(r!FolderID), "Nested folder label differs": r.MoveNext
    Loop
    r.Close
    Check PathError(999999, True) = PathError(999999, False), "Missing folder rejection differs"
    CurrentDb.Execute "UPDATE Folder SET ParentFolderID=FolderID WHERE FolderID=1", dbFailOnError
    Check PathError(1, False) = vbObjectError + 202 And PathError(1, True) = vbObjectError + 202, "Cycle/depth protection differs"
    CurrentDb.Execute "UPDATE Folder SET ParentFolderID=Null WHERE FolderID=1", dbFailOnError
    Set rows = PriorEvaluationITARows(1)
    For Each key In rows.Keys: Check rows(key) = PriorEvaluationITA(CLng(key)), "Prior NOB/incomplete/numeric average differs": Next
    Set rows = UXHistoryITARows(1)
    For Each key In rows.Keys: Check rows(key) = UXTableITA(CLng(key)), "Swift history ITA differs": Next
    rev = CurrentRevision(1): CompareCache 1, 0
    CurrentDb.Execute "UPDATE Report SET Trait33='1.0' WHERE ReportID=1", dbFailOnError
    CompareCache 1, vbObjectError + 8903
    CurrentDb.Execute "UPDATE Report SET Trait33='4.0' WHERE ReportID=1", dbFailOnError
    CurrentDb.Execute "UPDATE RevisionValue SET IsNullValue=Null WHERE RevisionID=" & rev & " AND FieldKey='Trait33'", dbFailOnError
    CompareCache 1, 0
    Check UXHistoryITARows(1)("1") = UXTableITA(1), "Nullable original-value flag differs"
    CurrentDb.Execute "UPDATE RevisionValue SET IsNullValue=False WHERE RevisionID=" & rev & " AND FieldKey='Trait33'", dbFailOnError
    CurrentDb.Execute "UPDATE Report SET CurrentRevisionID=0 WHERE ReportID=1", dbFailOnError
    CompareCache 1, vbObjectError + 8902
    CurrentDb.Execute "UPDATE Report SET CurrentRevisionID=" & rev & " WHERE ReportID=1", dbFailOnError
    CurrentDb.Execute "UPDATE RevisionDate SET DateState='Unknown',DateValue=Null WHERE RevisionID=" & rev & " AND FieldKey='DateReported'", dbFailOnError
    CompareCache 1, vbObjectError + 8911
    CurrentDb.Execute "UPDATE RevisionDate SET DateState='Date',DateValue=#2026-01-01# WHERE RevisionID=" & rev & " AND FieldKey='DateReported'", dbFailOnError
    CurrentDb.Execute "UPDATE RevisionChoice SET ChoiceValue='Yes' WHERE RevisionID=" & rev & " AND ChoiceKind='Retention'", dbFailOnError
    CompareCache 1, vbObjectError + 8912
    CurrentDb.Execute "UPDATE RevisionChoice SET ChoiceValue='Unspecified' WHERE RevisionID=" & rev & " AND ChoiceKind='Retention'", dbFailOnError
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RevisionValue WHERE RevisionID=" & rev & " AND FieldKey='Trait33'", dbOpenSnapshot)
    id = r!RevisionValueID: r.Close
    version = CLng(DLookup("FieldDefinitionVersion", "ReportRevision", "RevisionID=" & rev))
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionValueID=" & id, dbFailOnError
    CompareCache 1, vbObjectError + 8904
    CurrentDb.Execute "UPDATE ReportRevision SET FieldDefinitionVersion=0 WHERE RevisionID=" & rev, dbFailOnError
    CurrentDb.Execute "UPDATE Report SET Trait33=Null WHERE ReportID=1", dbFailOnError
    CompareCache 1, 0
    Check UXHistoryITARows(1)("1") = UXTableITA(1), "Missing saved trait was inferred"
    CurrentDb.Execute "INSERT INTO RevisionValue (RevisionValueID,RevisionID,FieldKey,ValueText,IsNullValue) VALUES (" & id & "," & rev & ",'Trait33','4.0',False)", dbFailOnError
    CurrentDb.Execute "UPDATE ReportRevision SET FieldDefinitionVersion=" & version & " WHERE RevisionID=" & rev, dbFailOnError
    CurrentDb.Execute "UPDATE Report SET Trait33='4.0' WHERE ReportID=1", dbFailOnError
    CompareCache 1, 0
    'All folder writes still pass through the original AssignFolder service.
    selection = "1,2": target = CreateFolder("SYNTHETIC destination")
    original = FolderBatchStamp(selection): count = DCount("*", "FolderMove")
    CurrentDb.TableDefs("FolderBatchChange").Fields("ReportID").ValidationRule = "<>2"
    On Error Resume Next
    operation = ApplyFolderBatch(selection, target, original): failure = Err.Number: Err.Clear
    On Error GoTo Failed
    CurrentDb.TableDefs("FolderBatchChange").Fields("ReportID").ValidationRule = ""
    Check failure <> 0 And FolderBatchStamp(selection) = original And DCount("*", "FolderMove") = count, "Late folder failure did not roll back"
    operation = ApplyFolderBatch(selection, target, original)
    Check UndoFolderBatch(operation), "Valid batch undo failed"
    Check DLookup("FolderID", "Report", "ReportID=1") = 3 And DLookup("FolderID", "Report", "ReportID=2") = 3, "Undo changed original assignments"
    operation = ApplyFolderBatch(selection, target, FolderBatchStamp(selection))
    AssignFolder 1, 3
    On Error Resume Next
    Call UndoFolderBatch(operation): failure = Err.Number: Err.Clear
    On Error GoTo Failed
    Check failure = vbObjectError + 1408 And DLookup("FolderID", "Report", "ReportID=2") = target, "Stale undo partially applied"
    CurrentDb.Execute "UPDATE Report SET WorkflowState='Approved' WHERE ReportID=2", dbFailOnError
    Check StampError("2", True) = StampError("2", False), "Locked-report preview rejection differs"
    TestBatchedReads = "PASS: " & checks & " batched read/guard, null/NOB, missing-field and rollback assertions": Exit Function
Failed:
    TestBatchedReads = "FAIL after " & checks & ": " & Err.Number & ": " & Err.Description
End Function
