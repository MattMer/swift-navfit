Attribute VB_Name = "ReportingSeniorHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 2450, , message
    checks = checks + 1
End Sub
Private Function Stored(ByVal id As Long, ByVal key As String) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT [" & key & "] FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Stored = r.Fields(0).Value: r.Close
End Function
Private Function SaveSenior(ByVal id As Long, ByVal kind As String, ByVal values As Variant) As Long
    SaveSenior = CommitDraft(id, kind, "Senior", "Synthetic", Null, Null, "", "Senior test", "Narrative", , , , , values)
End Function
Private Function Rejected(ByVal id As Long, ByVal values As Variant) As Boolean
    On Error GoTo Failed
    Call SaveSenior(id, "FITREP", values): Exit Function
Failed:
    Rejected = True
End Function
Private Function ReviewAcceptedAfterEdit(ByVal task As Long, ByVal revision As Long) As Boolean
    'F01: the assignee holds the report and may edit it before accepting.
    ReviewAcceptedAfterEdit = FinishReview(task, "Accepted", "Edited by the reviewer") And DLookup("FinishedRevisionID", "ReviewTask", "TaskID=" & task) = revision
End Function
Private Function RestoreRejected(ByVal id As Long, ByVal revision As Long) As Boolean
    On Error GoTo Failed
    Call RestoreDraft(id, revision): Exit Function
Failed:
    RestoreRejected = True
End Function
Public Function TestReportingSenior() As String
    Dim kind As Variant, values As Variant, keys As Variant, limits As Variant, bad As Variant
    Dim id As Long, first As Long, second As Long, copy As Long, task As Long, i As Long, before As Long, r As DAO.Recordset
    On Error GoTo Failed
    Call ConfigureWorkspace("Senior test", "Local reviewer")
    keys = ReportingSeniorKeys(): limits = Array(18, 5, 5, 14, 5)
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        values = Array("O'NEIL, A", "CAPT", "00110", "COMMANDER", "00001")
        id = SaveSenior(0, CStr(kind), values): first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check Stored(id, "ReportingSeniorName") = values(0) And Stored(id, "ReportingSeniorUIC") = "00001", "Text/zeros lost"
        Check DCount("*", "RevisionValue", "RevisionID=" & first) = UBound(DraftFieldKeys()) + 1, "Snapshot incomplete"
        Call CommitDraft(id, CStr(kind), "Senior", "Synthetic", Null, Null, "", "Omitted", "Narrative")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Omitted senior changed history"
        copy = DuplicateDraft(id): Check Stored(copy, "ReportingSeniorName") = values(0), "Duplicate lost senior"
        task = AssignReview(id, first, DLookup("ActiveProfileID", "Workspace"), Null, "Review")
        values(0) = "O'Neil, A": Call SaveSenior(id, CStr(kind), values)
        second = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check ReviewAcceptedAfterEdit(task, second), "The reviewer's own edit blocked accepting the review"
        Check RevisionFieldStatus(first, second, "ReportingSeniorName") = "Changed", "Case change invisible"
        Check OpenRevisionComparison(id, first), "Comparison failed"
        Forms!frmRevisionCompare!lstCompareFields = "ReportingSeniorName": Call LoadComparisonValues()
        Check Forms!frmRevisionCompare!txtOldValue = "O'NEIL, A" And Forms!frmRevisionCompare!txtNewValue = "O'Neil, A", "Comparison values wrong"
        DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
        Call RestoreDraft(id, first): Check Stored(id, "ReportingSeniorName") = "O'NEIL, A", "Restore lost senior"
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Check OpenReportingSenior(), "Panel failed to open"
        Check Forms!frmReportingSenior!txtReportingSeniorUIC = "00001", "Panel lost values"
        Forms!frmReportingSenior!txtReportingSeniorUIC = "99999"
        Check CancelReportingSenior(), "Cancel failed"
        Check Forms!frmEditor!txtReportingSeniorUIC = "00001", "Cancel changed pending values"
        Call OpenReportingSenior(): Forms!frmReportingSenior!txtReportingSeniorUIC = "00002"
        Check ApplyReportingSenior(), "Apply failed"
        Check Forms!frmEditor!txtReportingSeniorUIC = "00002" And Stored(id, "ReportingSeniorUIC") = "00001", "Apply committed report"
        Check DLookup("ReportingSeniorUIC", "RecoveryDraft", "ReportID=" & id) = "00002", "Apply did not checkpoint"
        Check SaveEditor(), "Editor save failed"
        Check Stored(id, "ReportingSeniorUIC") = "00002", "Save lost applied values"
        Call FixtureCloseEditor()
    Next
    For i = 0 To 4
        values = Array(Null, Null, Null, Null, Null): values(i) = String$(limits(i), "X")
        Call SaveSenior(id, "FITREP", values): Check Stored(id, CStr(keys(i))) = values(i), "Boundary truncated"
        values(i) = values(i) & "X": Check Rejected(id, values), "Oversized value accepted"
    Next
    before = DCount("*", "ReportRevision")
    For Each bad In Array(123, "Text" & vbTab, "Text" & vbCrLf)
        Check Rejected(id, Array(bad, Null, Null, Null, Null)), "Nontext/multiline accepted"
    Next
    Check Rejected(id, Array(Null)), "Malformed array accepted"
    Check DCount("*", "ReportRevision") = before, "Rejected save wrote history"
    values = Array(Null, "", Null, "", Null): Call SaveSenior(id, "FITREP", values)
    Check IsNull(Stored(id, "ReportingSeniorName")) And Not IsNull(Stored(id, "ReportingSeniorGrade")) And Len(Stored(id, "ReportingSeniorGrade")) = 0, "Null/empty collapsed"
    before = DCount("*", "ReportRevision")
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<> 'ReportingSeniorUIC'"
    Check Rejected(id, Array("NEW", "CAPT", "00110", "COMMANDER", "00003")), "Late write failure accepted"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = ""
    Check IsNull(Stored(id, "ReportingSeniorName")) And DCount("*", "ReportRevision") = before, "Late failure did not roll back"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Call OpenReportingSenior(): Forms!frmReportingSenior!txtReportingSeniorName = "PENDING NAME EXCEEDS LIMIT"
    Call ApplyReportingSenior(): Call FixtureCloseEditor()
    Check RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft")), "Recovery failed"
    Check Forms!frmEditor!txtReportingSeniorName = "PENDING NAME EXCEEDS LIMIT", "Raw pending value truncated"
    Check IsNull(Stored(id, "ReportingSeniorName")), "Recovery committed invalid name"
    Call FixtureCloseEditor()
    Call ArchiveDraft(id, True): Check Rejected(id, values), "Archived senior edit accepted"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Call OpenReportingSenior()
    Check Not Forms!frmReportingSenior!cmdApply.Enabled, "Archived panel can apply"
    Call CancelReportingSenior(): Call FixtureCloseEditor(): Call ArchiveDraft(id, False)
    first = DMin("RevisionID", "ReportRevision"): copy = DLookup("ReportID", "ReportRevision", "RevisionID=" & first)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & first & " AND FieldKey Like 'ReportingSenior*'", dbFailOnError
    CurrentDb.Execute "UPDATE ReportRevision SET SnapshotSchema=12 WHERE RevisionID=" & first, dbFailOnError
    Check OpenRevisionComparison(copy, first), "Legacy comparison failed"
    Forms!frmRevisionCompare!chkChangedOnly = False: Call RefreshComparisonFields()
    Forms!frmRevisionCompare!lstCompareFields = "ReportingSeniorName": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtOldValue = "[Not recorded in this revision]", "Legacy marker absent"
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
    Call RestoreDraft(copy, first): Check IsNull(Stored(copy, "ReportingSeniorName")), "Legacy restore invented senior"
    second = DLookup("CurrentRevisionID", "Report", "ReportID=" & copy)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & second & " AND FieldKey='ReportingSeniorUIC'", dbFailOnError
    Check RestoreRejected(copy, second), "Corrupt new snapshot restored"
    Set r = CurrentDb.OpenRecordset("RevisionValue", dbOpenDynaset)
    r.AddNew: r!RevisionID = second: r!FieldKey = "ReportingSeniorUIC": r!IsNullValue = True: r.Update: r.Close
    TestReportingSenior = "PASS: " & checks & " reporting senior assertions": Exit Function
Failed:
    TestReportingSenior = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function VerifyReportingSeniorRecovery() As Boolean
    Call RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft"))
    Call OpenReportingSenior()
    VerifyReportingSeniorRecovery = (Forms!frmReportingSenior!txtReportingSeniorName = "PENDING NAME EXCEEDS LIMIT")
    Call CancelReportingSenior(): Call FixtureCloseEditor()
End Function

Private Sub FixtureCloseEditor()
    ' Preserve the old abrupt-close semantics for persistence/recovery fixtures.
    ' Real save/discard/cancel exits are exercised by UXWorkflowHarness.
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Sub
