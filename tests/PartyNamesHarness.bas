Attribute VB_Name = "PartyNamesHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 3150, , message
    checks = checks + 1
End Sub
Private Function Seed() As Variant
    Seed = Array("MARY-JANE", "A.B", "O'NEIL", "JR", "SENIOR", "J", "K.L", "COUNSELOR", "M", "N.O")
End Function
Private Function Stored(ByVal id As Long, ByVal key As String) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT [" & key & "] FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Stored = r.Fields(0).Value: r.Close
End Function
Private Function SaveNames(ByVal id As Long, ByVal kind As String, ByVal values As Variant) As Long
    If id = 0 Then
        SaveNames = CommitDraft(id, kind, "Names", "Synthetic", Null, Null, "", "Names test", "Narrative", _
            administration:=Array("MEMBER, FULL", Null, Null, Null, Null), reportingSenior:=Array("SENIOR, F", Null, Null, Null, Null), _
            serviceDetails:=Array(Null, "FULL COUNSELOR", Null, Null, Null, Null), partyNames:=values)
    Else
        SaveNames = CommitDraft(id, kind, "Names", "Synthetic", Null, Null, "", "Names test", "Narrative", partyNames:=values)
    End If
End Function
Private Function Rejected(ByVal id As Long, ByVal values As Variant) As Boolean
    On Error GoTo Failed
    Call SaveNames(id, "FITREP", values): Exit Function
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
Public Function TestPartyNames() As String
    Dim kind As Variant, values As Variant, keys As Variant, key As Variant, bad As Variant, index As Variant
    Dim id As Long, first As Long, second As Long, copy As Long, task As Long, before As Long, i As Long, r As DAO.Recordset
    Dim limits As Variant
    On Error GoTo Failed
    Call ConfigureWorkspace("Names tests", "Synthetic reviewer")
    keys = PartyNamesKeys(): limits = Array(27, 3, 27, 10, 18, 1, 3, 20, 1, 3)
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        values = Seed(): id = SaveNames(0, CStr(kind), values): first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check DCount("*", "RevisionValue", "RevisionID=" & first) = UBound(DraftFieldKeys()) + 1, "Incomplete snapshot"
        For i = 0 To 9: Check Stored(id, CStr(keys(i))) = values(i), "Saved field changed": Next
        Call CommitDraft(id, CStr(kind), "Names", "Synthetic", Null, Null, "", "Omitted", "Narrative")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Omitted values wrote history"
        copy = DuplicateDraft(id)
        For i = 0 To 9: Check Stored(copy, CStr(keys(i))) = values(i), "Duplicate lost field": Next
        For Each index In Array(0, 4, 7)
            first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            task = AssignReview(id, first, DLookup("ActiveProfileID", "Workspace"), Null, "Feature review")
            Select Case index
                Case 0: values(0) = "CHANGED MEMBER"
                Case 4: values(4) = "CHANGED SENIOR"
                Case 7: values(7) = "CHANGED COUNSELOR"
            End Select
            Call SaveNames(id, CStr(kind), values): second = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            Check ReviewAcceptedAfterEdit(task, second), "The reviewer's own edit blocked accepting the review"
            Check RevisionFieldStatus(first, second, CStr(keys(index))) = "Changed", "Feature change invisible"
            Check OpenRevisionComparison(id, first), "Comparison failed"
            Forms!frmRevisionCompare!lstCompareFields = keys(index): Call LoadComparisonValues()
            Check Forms!frmRevisionCompare!txtNewValue = values(index), "Comparison value wrong"
            DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
        Next
        first = DMin("RevisionID", "ReportRevision", "ReportID=" & id): Call RestoreDraft(id, first): values = Seed()
        For i = 0 To 9: Check Stored(id, CStr(keys(i))) = values(i), "Restore lost field": Next
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Check OpenPartyNames(), "Panel failed"
        For i = 0 To 9: Check Forms!frmPartyNames.Controls("txt" & keys(i)) = values(i), "Panel lost field": Next
        Forms!frmPartyNames!txtMemberFirstName = "CANCEL ME": Call CancelPartyNames()
        Check Forms!frmEditor!txtMemberFirstName = "MARY-JANE", "Cancel changed pending values"
        Call OpenPartyNames(): Forms!frmPartyNames!txtReportingSeniorLastName = "NEW SENIOR"
        Forms!frmPartyNames!txtCounselorLastName = "NEW COUNSELOR": Forms!frmPartyNames!txtMemberFirstName = "NEW MEMBER"
        Check ApplyPartyNames(), "Apply failed"
        Check Stored(id, "MemberFirstName") = "MARY-JANE" And Forms!frmEditor!txtMemberFirstName = "NEW MEMBER", "Apply committed revision"
        Check DLookup("CounselorLastName", "RecoveryDraft", "ReportID=" & id) = "NEW COUNSELOR", "Apply did not checkpoint"
        Check SaveEditor(), "Editor save failed"
        Check Stored(id, "ReportingSeniorLastName") = "NEW SENIOR" And Stored(id, "CounselorLastName") = "NEW COUNSELOR" And Stored(id, "MemberFirstName") = "NEW MEMBER", "Editor save lost features"
        Check Stored(id, "OfficialName") = "MEMBER, FULL" And Stored(id, "ReportingSeniorName") = "SENIOR, F" And Stored(id, "CounselorName") = "FULL COUNSELOR", "Name parts replaced full names"
        Call FixtureCloseEditor()
    Next
    For Each index In Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9)
        values = Seed(): values(index) = String$(limits(index), "X")
        Call SaveNames(id, "FITREP", values): Check Stored(id, CStr(keys(index))) = values(index), "Boundary truncated"
        values(index) = values(index) & "X": Check Rejected(id, values), "Overlength accepted"
    Next
    before = DCount("*", "ReportRevision")
    For i = 0 To 9
        values = Seed(): values(i) = 123: Check Rejected(id, values), "Nontext accepted"
        values(i) = "X" & vbTab: Check Rejected(id, values), "Tab accepted"
        values(i) = "X" & vbCrLf: Check Rejected(id, values), "Multiline accepted"
    Next
    Check Rejected(id, Array(Null)), "Malformed array accepted"
    Check DCount("*", "ReportRevision") = before, "Rejected save wrote history"
    values = Seed(): values(0) = "MARY" & ChrW(&HE9): Call SaveNames(id, "FITREP", values)
    Check Stored(id, "MemberFirstName") = values(0), "Accented name changed"
    first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
    values(0) = LCase$(values(0)): Call SaveNames(id, "FITREP", values)
    Check Stored(id, "MemberFirstName") = values(0), "Case normalized"
    second = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
    Check second <> first And RevisionFieldStatus(first, second, "MemberFirstName") = "Changed", "Case change hidden"
    For i = 0 To 9
        values(i) = "": Call SaveNames(id, "FITREP", values)
        Check Not IsNull(Stored(id, CStr(keys(i)))) And Stored(id, CStr(keys(i))) = "", "Empty entry collapsed"
        values(i) = Null: Call SaveNames(id, "FITREP", values)
        Check IsNull(Stored(id, CStr(keys(i)))), "Null entry collapsed"
    Next
    values = Seed(): Call SaveNames(id, "FITREP", values)
    before = DCount("*", "ReportRevision")
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<> 'CounselorMiddleInitial'"
    values = Seed(): values(0) = "ROLLBACK MEMBER"
    Check Rejected(id, values), "Late write failure accepted"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = ""
    Check Stored(id, "MemberFirstName") = "MARY-JANE" And Stored(id, "ReportingSeniorLastName") = "SENIOR" And DCount("*", "ReportRevision") = before, "Late failure did not roll back"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Call OpenPartyNames()
    Forms!frmPartyNames!txtReportingSeniorLastName = "PENDING OVERLONG SENIOR LAST NAME"
    Forms!frmPartyNames!txtCounselorLastName = String$(45, "X")
    Forms!frmPartyNames!txtCounselorFirstInitial = "AB"
    Call ApplyPartyNames(): Call FixtureCloseEditor()
    Check RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft")), "Recovery failed"
    Check Forms!frmEditor!txtReportingSeniorLastName = "PENDING OVERLONG SENIOR LAST NAME" And Forms!frmEditor!txtCounselorLastName = String$(45, "X") And Forms!frmEditor!txtCounselorFirstInitial = "AB", "Raw recovery truncated"
    Check Stored(id, "ReportingSeniorLastName") = "SENIOR" And Stored(id, "CounselorFirstInitial") = "M", "Recovery auto committed"
    Call FixtureCloseEditor()
    Call ArchiveDraft(id, True): Check Rejected(id, Seed()), "Archive guard absent"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Call OpenPartyNames()
    Check Not Forms!frmPartyNames!cmdApply.Enabled, "Archived panel can apply"
    Call CancelPartyNames(): Call FixtureCloseEditor(): Call ArchiveDraft(id, False)
    first = DMin("RevisionID", "ReportRevision"): copy = DLookup("ReportID", "ReportRevision", "RevisionID=" & first)
    For Each key In keys: CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & first & " AND FieldKey='" & key & "'", dbFailOnError: Next
    CurrentDb.Execute "UPDATE ReportRevision SET SnapshotSchema=18 WHERE RevisionID=" & first, dbFailOnError
    Check OpenRevisionComparison(copy, first), "Legacy comparison failed"
    Forms!frmRevisionCompare!chkChangedOnly = False: Call RefreshComparisonFields()
    Forms!frmRevisionCompare!lstCompareFields = "CounselorMiddleInitial": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtOldValue = "[Not recorded in this revision]", "Legacy marker absent"
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo: Call RestoreDraft(copy, first)
    For Each key In keys: Check IsNull(Stored(copy, CStr(key))), "Legacy restore invented details": Next
    second = DLookup("CurrentRevisionID", "Report", "ReportID=" & copy)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & second & " AND FieldKey='CounselorMiddleInitial'", dbFailOnError
    Check RestoreRejected(copy, second), "Corrupt new snapshot restored"
    Set r = CurrentDb.OpenRecordset("RevisionValue", dbOpenDynaset)
    r.AddNew: r!RevisionID = second: r!FieldKey = "CounselorMiddleInitial": r!IsNullValue = True: r.Update: r.Close
    TestPartyNames = "PASS: " & checks & " structured name assertions": Exit Function
Failed:
    TestPartyNames = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function VerifyPartyNamesRecovery() As Boolean
    Call RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft"))
    VerifyPartyNamesRecovery = (Forms!frmEditor!txtReportingSeniorLastName = "PENDING OVERLONG SENIOR LAST NAME" And Forms!frmEditor!txtCounselorLastName = String$(45, "X") And Forms!frmEditor!txtCounselorFirstInitial = "AB")
    Call FixtureCloseEditor()
End Function

Private Sub FixtureCloseEditor()
    ' Preserve the old abrupt-close semantics for persistence/recovery fixtures.
    ' Real save/discard/cancel exits are exercised by UXWorkflowHarness.
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Sub
