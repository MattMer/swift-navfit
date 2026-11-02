Attribute VB_Name = "ReportPeopleHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 2950, , message
    checks = checks + 1
End Sub
Private Function Seed() As Variant
    Seed = Array("000123456", "0012345678", "123 SYNTHETIC ST", "SUITE A", "TEST CITY", "CA", "000001234", "0001234567", "0001234", "O'NEIL, A", "2024-02-29", "SYNTHETIC SENIOR", "2026-10-03")
End Function
Private Function Stored(ByVal id As Long, ByVal key As String) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT [" & key & "] FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Stored = r.Fields(0).Value: r.Close
End Function
Private Function SavePeople(ByVal id As Long, ByVal kind As String, ByVal values As Variant) As Long
    SavePeople = CommitDraft(id, kind, "People", "Synthetic", Null, Null, "", "People test", "Narrative", , , , , , , , , , values)
End Function
Private Function Rejected(ByVal id As Long, ByVal values As Variant) As Boolean
    On Error GoTo Failed
    Call SavePeople(id, "FITREP", values): Exit Function
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
Public Function TestReportPeople() As String
    Dim kind As Variant, values As Variant, keys As Variant, key As Variant, bad As Variant, index As Variant
    Dim id As Long, first As Long, second As Long, copy As Long, task As Long, before As Long, i As Long, r As DAO.Recordset
    Dim limits As Variant
    On Error GoTo Failed
    Call ConfigureWorkspace("People tests", "Synthetic reviewer")
    keys = ReportPeopleKeys(): limits = Array(9, 10, 30, 30, 15, 2, 9, 10, 7, 28, 10, 28, 10)
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        values = Seed(): id = SavePeople(0, CStr(kind), values): first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check DCount("*", "RevisionValue", "RevisionID=" & first) = UBound(DraftFieldKeys()) + 1, "Incomplete snapshot"
        For i = 0 To 12: Check Stored(id, CStr(keys(i))) = values(i), "Saved field changed": Next
        Call CommitDraft(id, CStr(kind), "People", "Synthetic", Null, Null, "", "Omitted", "Narrative")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Omitted values wrote history"
        copy = DuplicateDraft(id)
        For i = 0 To 12: Check Stored(copy, CStr(keys(i))) = values(i), "Duplicate lost field": Next
        For Each index In Array(0, 2, 9)
            first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            task = AssignReview(id, first, DLookup("ActiveProfileID", "Workspace"), Null, "Feature review")
            Select Case index
                Case 0: values(0) = "000123457"
                Case 2: values(2) = "CHANGED ADDRESS"
                Case 9: values(9) = "CHANGED RATER"
            End Select
            Call SavePeople(id, CStr(kind), values): second = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            Check ReviewAcceptedAfterEdit(task, second), "The reviewer's own edit blocked accepting the review"
            Check RevisionFieldStatus(first, second, CStr(keys(index))) = "Changed", "Feature change invisible"
            Check OpenRevisionComparison(id, first), "Comparison failed"
            Forms!frmRevisionCompare!lstCompareFields = keys(index): Call LoadComparisonValues()
            Check Forms!frmRevisionCompare!txtNewValue = values(index), "Comparison value wrong"
            DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
        Next
        first = DMin("RevisionID", "ReportRevision", "ReportID=" & id): Call RestoreDraft(id, first): values = Seed()
        For i = 0 To 12: Check Stored(id, CStr(keys(i))) = values(i), "Restore lost field": Next
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Check OpenReportPeople(), "Panel failed"
        For i = 0 To 12: Check Forms!frmReportPeople.Controls("txt" & keys(i)) = values(i), "Panel lost field": Next
        Forms!frmReportPeople!txtRaterName = "CANCEL ME": Call CancelReportPeople()
        Check Forms!frmEditor!txtRaterName = "O'NEIL, A", "Cancel changed pending values"
        Call OpenReportPeople(): Forms!frmReportPeople!txtReportingSeniorSSN = "000123458"
        Forms!frmReportPeople!txtReportingSeniorAddress1 = "NEW ADDRESS": Forms!frmReportPeople!txtRaterName = "NEW RATER"
        Check ApplyReportPeople(), "Apply failed"
        Check Stored(id, "RaterName") = "O'NEIL, A" And Forms!frmEditor!txtRaterName = "NEW RATER", "Apply committed revision"
        Check DLookup("ReportingSeniorAddress1", "RecoveryDraft", "ReportID=" & id) = "NEW ADDRESS", "Apply did not checkpoint"
        Check SaveEditor(), "Editor save failed"
        Check Stored(id, "ReportingSeniorSSN") = "000123458" And Stored(id, "ReportingSeniorAddress1") = "NEW ADDRESS" And Stored(id, "RaterName") = "NEW RATER", "Editor save lost features"
        Check IsNull(Stored(id, "MemberSSN")) And IsNull(Stored(id, "ReportingSeniorName")), "People details inferred other identities"
        Call FixtureCloseEditor()
    Next
    For Each index In Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 11)
        values = Seed(): If index < 2 Then values(index) = String$(limits(index), "0") Else values(index) = String$(limits(index), "X")
        Call SavePeople(id, "FITREP", values): Check Stored(id, CStr(keys(index))) = values(index), "Boundary truncated"
        values(index) = values(index) & "X": Check Rejected(id, values), "Overlength accepted"
    Next
    before = DCount("*", "ReportRevision")
    For i = 0 To 12
        values = Seed(): values(i) = 123: Check Rejected(id, values), "Nontext accepted"
        values(i) = "X" & vbTab: Check Rejected(id, values), "Tab accepted"
        values(i) = "X" & vbCrLf: Check Rejected(id, values), "Multiline accepted"
    Next
    For Each index In Array(0, 1)
        For Each bad In Array("12A", " 123", "123 ", "123-45", "123.45", "1E3")
            values = Seed(): values(index) = bad: Check Rejected(id, values), "Nondigit identifier accepted"
        Next
    Next
    For Each index In Array(10, 12)
        For Each bad In Array("2023-02-29", "1900-02-29", "2024-02-30", "2026-13-01", "2026-10-00", "10/03/2026", "2026-10-03T12:00", "2026-10-03 ")
            values = Seed(): values(index) = bad: Check Rejected(id, values), "Invalid date accepted"
        Next
    Next
    Check Rejected(id, Array(Null)), "Malformed array accepted"
    Check DCount("*", "ReportRevision") = before, "Rejected save wrote history"
    values = Seed(): values(10) = "2000-02-29": values(12) = "2000-02-29": Call SavePeople(id, "FITREP", values)
    Check Stored(id, "RaterDate") = "2000-02-29" And Stored(id, "SeniorRaterDate") = "2000-02-29", "Century leap date rejected"
    For i = 0 To 12
        values(i) = "": Call SavePeople(id, "FITREP", values)
        Check Not IsNull(Stored(id, CStr(keys(i)))) And Stored(id, CStr(keys(i))) = "", "Empty entry collapsed"
        values(i) = Null: Call SavePeople(id, "FITREP", values)
        Check IsNull(Stored(id, CStr(keys(i)))), "Null entry collapsed"
    Next
    values = Seed(): values(0) = "0": values(1) = "00": Call SavePeople(id, "FITREP", values)
    Check Stored(id, "ReportingSeniorSSN") = "0" And Stored(id, "ReportingSeniorDoDID") = "00", "Partial identifiers rejected"
    before = DCount("*", "ReportRevision")
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<> 'SeniorRaterDate'"
    Check Rejected(id, Seed()), "Late write failure accepted"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = ""
    Check Stored(id, "ReportingSeniorSSN") = "0" And DCount("*", "ReportRevision") = before, "Late failure did not roll back"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Call OpenReportPeople()
    Forms!frmReportPeople!txtReportingSeniorSSN = "PENDING INVALID ID"
    Forms!frmReportPeople!txtReportingSeniorAddress1 = String$(45, "X")
    Forms!frmReportPeople!txtSeniorRaterDate = "PENDING INVALID DATE"
    Call ApplyReportPeople(): Call FixtureCloseEditor()
    Check RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft")), "Recovery failed"
    Check Forms!frmEditor!txtReportingSeniorSSN = "PENDING INVALID ID" And Forms!frmEditor!txtReportingSeniorAddress1 = String$(45, "X") And Forms!frmEditor!txtSeniorRaterDate = "PENDING INVALID DATE", "Raw recovery truncated"
    Check Stored(id, "ReportingSeniorSSN") = "0" And Stored(id, "SeniorRaterDate") = "2026-10-03", "Recovery auto committed"
    Call FixtureCloseEditor()
    Call ArchiveDraft(id, True): Check Rejected(id, Seed()), "Archive guard absent"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Call OpenReportPeople()
    Check Not Forms!frmReportPeople!cmdApply.Enabled, "Archived panel can apply"
    Call CancelReportPeople(): Call FixtureCloseEditor(): Call ArchiveDraft(id, False)
    first = DMin("RevisionID", "ReportRevision"): copy = DLookup("ReportID", "ReportRevision", "RevisionID=" & first)
    For Each key In keys: CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & first & " AND FieldKey='" & key & "'", dbFailOnError: Next
    CurrentDb.Execute "UPDATE ReportRevision SET SnapshotSchema=17 WHERE RevisionID=" & first, dbFailOnError
    Check OpenRevisionComparison(copy, first), "Legacy comparison failed"
    Forms!frmRevisionCompare!chkChangedOnly = False: Call RefreshComparisonFields()
    Forms!frmRevisionCompare!lstCompareFields = "SeniorRaterDate": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtOldValue = "[Not recorded in this revision]", "Legacy marker absent"
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo: Call RestoreDraft(copy, first)
    For Each key In keys: Check IsNull(Stored(copy, CStr(key))), "Legacy restore invented details": Next
    second = DLookup("CurrentRevisionID", "Report", "ReportID=" & copy)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & second & " AND FieldKey='SeniorRaterDate'", dbFailOnError
    Check RestoreRejected(copy, second), "Corrupt new snapshot restored"
    Set r = CurrentDb.OpenRecordset("RevisionValue", dbOpenDynaset)
    r.AddNew: r!RevisionID = second: r!FieldKey = "SeniorRaterDate": r!IsNullValue = True: r.Update: r.Close
    TestReportPeople = "PASS: " & checks & " reporting contact and rater assertions": Exit Function
Failed:
    TestReportPeople = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function VerifyReportPeopleRecovery() As Boolean
    Call RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft"))
    VerifyReportPeopleRecovery = (Forms!frmEditor!txtReportingSeniorSSN = "PENDING INVALID ID" And Forms!frmEditor!txtReportingSeniorAddress1 = String$(45, "X") And Forms!frmEditor!txtSeniorRaterDate = "PENDING INVALID DATE")
    Call FixtureCloseEditor()
End Function

Private Sub FixtureCloseEditor()
    ' Preserve the old abrupt-close semantics for persistence/recovery fixtures.
    ' Real save/discard/cancel exits are exercised by UXWorkflowHarness.
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Sub
