Attribute VB_Name = "PersistenceHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 8950, , message
    n = n + 1
End Sub
Private Function RejectSQL(ByVal sql As String) As Boolean
    On Error GoTo Rejected
    CurrentDb.Execute sql, dbFailOnError
    Exit Function
Rejected:
    RejectSQL = True
End Function
Public Function TestAdditionalField() As String
    Dim id As Long, rev As Long, f As Object, checkpoint As Long, stage As String
    On Error GoTo Failed
    stage = "Configure"
    ConfigureWorkspace "SYNTHETIC CATALOGUE EXTENSION", "Tester"
    stage = "Read older catalogue snapshot"
    id = DLookup("ReportID", "Report", "Title='Catalogue extension'")
    rev = CurrentRevision(id)
    If DLookup("FieldDefinitionVersion", "ReportRevision", "RevisionID=" & rev) <> 1 Then Err.Raise 5, , "Prior snapshot definition was relabeled"
    If MajorRevisionHasField(rev, "CatalogueTestValue") Or MajorRevisionHasField(rev, "CatalogueTestDate") Then Err.Raise 5, , "Prior snapshot was rewritten for new fields"
    Set f = CreateObject("Scripting.Dictionary"): f.Add "CatalogueTestValue", "Saved through catalogue"
    f.Add "CatalogueTestDate", #2/29/2024#
    stage = "Save appended values"
    CommitSavedFields id, rev, f, "Catalogue extension"
    If MajorRevisionValue(CurrentRevision(id), "CatalogueTestValue") <> "Saved through catalogue" Then Err.Raise 5, , "New field did not save"
    If MajorRevisionValue(CurrentRevision(id), "CatalogueTestDate") <> "2024-02-29" Then Err.Raise 5, , "New date snapshot is not invariant ISO"
    stage = "Restore"
    RestoreDraft id, rev
    If DLookup("FieldDefinitionVersion", "ReportRevision", "RevisionID=" & rev) <> 1 Then Err.Raise 5, , "Restore relabeled the source snapshot"
    If Not IsNull(MajorRevisionValue(CurrentRevision(id), "CatalogueTestValue")) Then Err.Raise 5, , "New field did not restore"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    stage = "Pending buffer"
    Forms!frmEditor!txtCatalogueTestValue = "Raw pending catalogue value"
    Forms!frmEditor!txtCatalogueTestDate = "Raw pending date"
    stage = "Checkpoint"
    If Not CheckpointEditor() Then Err.Raise 5, , "New field did not checkpoint"
    If DLookup("CatalogueTestValue", "RecoveryDraft") <> "Raw pending catalogue value" Then Err.Raise 5, , "Checkpoint omitted new field"
    If DLookup("CatalogueTestDate", "RecoveryDraft") <> "Raw pending date" Then Err.Raise 5, , "Checkpoint coerced raw date entry"
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
    TestAdditionalField = "PASS: metadata-only synthetic text/date fields saved, snapshotted, restored and checkpointed": Exit Function
Failed:
    TestAdditionalField = "FAIL: " & stage & ": " & Err.Source & ": " & Err.Number & " " & Err.Description
End Function

Public Function VerifyAdditionalRecovery() As Boolean
    Dim id As Long
    BindActiveContributor CLng(DLookup("UserProfileID", "UserProfile"))
    id = DLookup("RecoveryDraftID", "RecoveryDraft")
    If Not RestoreRecoveryDraft(id, False) Then Exit Function
    VerifyAdditionalRecovery = (Forms!frmEditor!txtCatalogueTestValue = "Raw pending catalogue value" And Forms!frmEditor!txtCatalogueTestDate = "Raw pending date")
    UXLeaveEditor "Discard"
End Function
Private Function RejectRead(ByVal id As Long) As Boolean
    Dim f As Object
    On Error GoTo Rejected
    Set f = SavedFields(id)
    Exit Function
Rejected:
    RejectRead = True
End Function
Private Function RejectSave(ByVal id As Long, ByVal f As Object) As Boolean
    On Error GoTo Rejected
    CommitSavedFields id, CurrentRevision(id), f, "Expected failure"
    Exit Function
Rejected:
    RejectSave = True
End Function
Public Function TestPersistence() As String
    Dim a As Long, b As Long, rev As Long, old As Long, count As Long, f As Object, saved As Object, key As Variant, r As DAO.Recordset
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC PERSISTENCE", "Tester"
    a = CommitDraft(0, "EVAL", "First", "Synthetic", #1/1/2026#, #9/30/2026#, "notes")
    b = CommitDraft(0, "FITREP", "Second", "Synthetic", Null, Null, "")
    rev = CurrentRevision(a)
    Check rev = DLookup("CurrentRevisionID", "Report", "ReportID=" & a), "Head/cache pointer differs"
    Check DCount("*", "RevisionDate", "RevisionID=" & rev & " AND DateState='Missing'") = 5, "Null date facts were not retained"
    Check DCount("*", "RevisionChoice", "RevisionID=" & rev & " AND ChoiceValue='Unspecified'") = 2, "Missing choices were not represented"
    Check RejectSQL("UPDATE ReportHead SET RevisionID=" & CurrentRevision(b) & " WHERE ReportID=" & a), "Cross-report revision accepted"
    Check RejectSQL("UPDATE Report SET FormKind='OTHER' WHERE ReportID=" & a), "Unsupported new report kind accepted"
    Check RejectSQL("UPDATE Report SET WorkflowState='OTHER' WHERE ReportID=" & a), "Unsupported new workflow accepted"
    Check RejectSQL("UPDATE RevisionDate SET DateState='Date' WHERE RevisionID=" & rev), "Date state without date accepted"
    Check RejectSQL("UPDATE RevisionChoice SET ChoiceValue='OTHER' WHERE RevisionID=" & rev), "Unsupported choice accepted"
    CurrentDb.Execute "UPDATE RevisionChoice SET ChoiceValue='Yes' WHERE RevisionID=" & rev & " AND ChoiceKind='Retention'", dbFailOnError
    Check RejectRead(a), "Independent typed choice write became authoritative"
    CurrentDb.Execute "UPDATE RevisionChoice SET ChoiceValue='Unspecified' WHERE RevisionID=" & rev & " AND ChoiceKind='Retention'", dbFailOnError
    CommitReportChoice a, rev, "Retention", "No": rev = CurrentRevision(a)
    Check MajorRevisionValue(rev, "RetentionYes") = "No" And MajorRevisionValue(rev, "RetentionNo") = "Yes", "Single choice did not produce lossless native flags"
    Set f = CreateObject("Scripting.Dictionary")
    f.Add "CounselingDate", "NOT PERF": f.Add "CounselorName", "Required explanation"
    f.Add "DateReported", "2024-02-29": f.Add "ReadinessDate", "": f.Add "RaterDate", "2026-10-22"
    f.Add "RetentionYes", "Yes": f.Add "RetentionNo", "Yes"
    CommitSavedFields a, rev, f, "Preserve conflict for review"
    old = rev: rev = CurrentRevision(a)
    Check DLookup("DateState", "RevisionDate", "RevisionID=" & rev & " AND FieldKey='CounselingDate'") = "NOT PERF", "Counseling status lost"
    Check DLookup("DateValue", "RevisionDate", "RevisionID=" & rev & " AND FieldKey='DateReported'") = #2/29/2024#, "Leap date lost"
    Check DLookup("DateState", "RevisionDate", "RevisionID=" & rev & " AND FieldKey='ReadinessDate'") = "Blank", "Empty date coerced to Null"
    Check DLookup("ChoiceValue", "RevisionChoice", "RevisionID=" & rev & " AND ChoiceKind='Retention'") = "Conflict", "Conflicting flags silently coerced"
    Check MajorRevisionValue(rev, "CounselorName") = "Required explanation", "Counseling explanation lost"
    RestoreDraft a, old
    Check DLookup("DateState", "RevisionDate", "RevisionID=" & CurrentRevision(a) & " AND FieldKey='CounselingDate'") = "Missing", "Restore retained newer typed facts"
    RestoreDraft a, rev
    Check MajorRevisionValue(CurrentRevision(a), "CounselingDate") = "NOT PERF", "Restore lost status snapshot"
    Set saved = SavedFields(a)
    CurrentDb.Execute "UPDATE Report SET Title='Tampered projection' WHERE ReportID=" & a, dbFailOnError
    Check RejectRead(a), "Independent cache write became authoritative"
    Check MajorRevisionValue(CurrentRevision(a), "Title") = "First", "Cache tamper changed snapshot"
    CurrentDb.Execute "UPDATE Report SET Title='First' WHERE ReportID=" & a, dbFailOnError
    AssertReportCache a
    SetFieldDefault "Report", a, "CounselingDate", "NOT REQ"
    Check DCount("*", "FieldDefaultLink", "ReportID=" & a & " AND LinkState='Verified'") = 1, "Default relationship missing"
    Check RejectSQL("DELETE FROM Report WHERE ReportID=" & a), "Referenced report deletion accepted"
    Check RejectSQL("UPDATE FieldDefaultLink SET GroupID=999999 WHERE ReportID=" & a), "Invalid/multiple default targets accepted"
    rev = CurrentRevision(a): count = DCount("*", "ReportRevision")
    f.RemoveAll: f.Add "Title", "Must roll back"
    CurrentDb.TableDefs("RevisionDate").Fields("DateState").ValidationRule = "='Reject all facts'"
    Check RejectSave(a, f), "Injected typed fact failure was accepted"
    CurrentDb.TableDefs("RevisionDate").Fields("DateState").ValidationRule = ""
    Check CurrentRevision(a) = rev And DCount("*", "ReportRevision") = count, "Typed failure retained revision/head"
    Check DLookup("Title", "Report", "ReportID=" & a) = "First", "Typed failure retained cache write"
    'Synthetic legacy values: preserve exactly, derive Unknown, and expose review state.
    CurrentDb.Execute "UPDATE Report SET DateReported='LEGACY' WHERE ReportID=" & a, dbFailOnError
    CurrentDb.Execute "UPDATE RevisionValue SET ValueText='LEGACY',IsNullValue=False WHERE RevisionID=" & rev & " AND FieldKey='DateReported'", dbFailOnError
    CurrentDb.Execute "DELETE FROM RevisionDate WHERE RevisionID=" & rev & " AND FieldKey='DateReported'", dbFailOnError
    WriteRevisionFacts rev
    Check DLookup("DateState", "RevisionDate", "RevisionID=" & rev & " AND FieldKey='DateReported'") = "Unknown", "Legacy date silently parsed/coerced"
    Set saved = SavedFields(a): Check saved("DateReported") = "LEGACY", "Unknown raw date lost"
    count = DCount("*", "RevisionDate"): old = DCount("*", "ReportHead")
    MigrateReportState: MigrateReportState
    Check DCount("*", "RevisionDate") = count And DCount("*", "ReportHead") = old, "State migration is not idempotent"
    For Each key In Array("EVAL", "CHIEFEVAL", "FITREP")
        Check NAVFITMapRows(CStr(key)).Count > 60 And PDFMapRows(CStr(key)).Count > 35, "Canonical adapter coverage missing"
        Check ReportPageForField(CStr(key), "Trait39") = 0, "Seven-trait Front Page geometry changed"
    Next
    TestPersistence = "PASS: " & n & " persistence, typed facts, constraints, rollback and legacy-preservation assertions": Exit Function
Failed:
    TestPersistence = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
