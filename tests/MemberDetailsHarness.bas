Attribute VB_Name = "MemberDetailsHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 2650, , message
    checks = checks + 1
End Sub
Private Function Seed() As Variant
    Seed = Array("000123456", "0001234567", "Yes", "No", "No", Null, "2024-02-29")
End Function
Private Function Stored(ByVal id As Long, ByVal key As String) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT [" & key & "] FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Stored = r.Fields(0).Value: r.Close
End Function
Private Function SaveMember(ByVal id As Long, ByVal kind As String, ByVal values As Variant) As Long
    SaveMember = CommitDraft(id, kind, "Member", "Synthetic", Null, Null, "", "Member test", "Narrative", , , , , , , values)
End Function
Private Function Rejected(ByVal id As Long, ByVal values As Variant) As Boolean
    On Error GoTo Failed
    Call SaveMember(id, "FITREP", values): Exit Function
Failed:
    Rejected = True
End Function
Private Function ReviewRejected(ByVal task As Long) As Boolean
    On Error GoTo Failed
    Call FinishReview(task, "Accepted", ""): Exit Function
Failed:
    ReviewRejected = True
End Function
Private Function RestoreRejected(ByVal id As Long, ByVal revision As Long) As Boolean
    On Error GoTo Failed
    Call RestoreDraft(id, revision): Exit Function
Failed:
    RestoreRejected = True
End Function
Public Function TestMemberDetails() As String
    Dim kind As Variant, values As Variant, keys As Variant, key As Variant, bad As Variant, index As Variant
    Dim id As Long, first As Long, second As Long, copy As Long, task As Long, before As Long, i As Long, r As DAO.Recordset
    On Error GoTo Failed
    Call ConfigureWorkspace("Member tests", "Synthetic reviewer"): keys = MemberDetailsKeys()
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        values = Seed(): id = SaveMember(0, CStr(kind), values): first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check DCount("*", "RevisionValue", "RevisionID=" & first) = UBound(DraftFieldKeys()) + 1, "Incomplete snapshot"
        Check Stored(id, "MemberSSN") = "000123456" And Stored(id, "MemberDoDID") = "0001234567", "Leading zeros lost"
        Check Stored(id, "ComponentActive") = "Yes" And Stored(id, "ComponentTAR") = "No" And IsNull(Stored(id, "ComponentATADSW")), "Component states collapsed"
        Check Stored(id, "DateReported") = "2024-02-29", "Leap date changed"
        Call CommitDraft(id, CStr(kind), "Member", "Synthetic", Null, Null, "", "Omitted", "Narrative")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Omitted details wrote history"
        copy = DuplicateDraft(id)
        Check Stored(copy, "MemberSSN") = values(0) And Stored(copy, "ComponentActive") = "Yes" And Stored(copy, "DateReported") = values(6), "Duplicate lost features"
        For Each index In Array(0, 2, 6)
            first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            task = AssignReview(id, first, DLookup("ActiveProfileID", "Workspace"), Null, "Feature review")
            Select Case index
                Case 0: values(0) = "000000001"
                Case 2: values(2) = "No"
                Case 6: values(6) = "2026-10-03"
            End Select
            Call SaveMember(id, CStr(kind), values): second = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            Check ReviewRejected(task), "Feature edit left review current"
            Call FinishReview(task, "Cancelled", "Changed details")
            Check RevisionFieldStatus(first, second, CStr(keys(index))) = "Changed", "Feature change invisible"
            Check OpenRevisionComparison(id, first), "Comparison failed"
            Forms!frmRevisionCompare!lstCompareFields = keys(index): Call LoadComparisonValues()
            Check Forms!frmRevisionCompare!txtNewValue = values(index), "Comparison value wrong"
            DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
        Next
        first = DMin("RevisionID", "ReportRevision", "ReportID=" & id): Call RestoreDraft(id, first)
        Check Stored(id, "MemberSSN") = "000123456" And Stored(id, "ComponentActive") = "Yes" And Stored(id, "DateReported") = "2024-02-29", "Restore lost features"
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Check OpenMemberDetails(), "Panel failed"
        Check Forms!frmMemberDetails!txtMemberDoDID = "0001234567", "Panel lost identifier"
        Forms!frmMemberDetails!txtMemberSSN = "999999999": Call CancelMemberDetails()
        Check Forms!frmEditor!txtMemberSSN = "000123456", "Cancel changed pending values"
        Call OpenMemberDetails(): Forms!frmMemberDetails!txtDateReported = "2026-10-03"
        Check ApplyMemberDetails(), "Apply failed"
        Check Forms!frmEditor!txtDateReported = "2026-10-03" And Stored(id, "DateReported") = "2024-02-29", "Apply committed revision"
        Check DLookup("DateReported", "RecoveryDraft", "ReportID=" & id) = "2026-10-03", "Apply failed to checkpoint"
        Check SaveEditor(), "Editor save failed"
        Check Stored(id, "DateReported") = "2026-10-03", "Editor save lost date"
        Call FixtureCloseEditor()
    Next
    before = DCount("*", "ReportRevision")
    For Each bad In Array("123-45-6789", "ABCDEFGHI", "1234567890", " 123", 123)
        values = Seed(): values(0) = bad: Check Rejected(id, values), "Invalid identifier accepted"
    Next
    values = Seed(): values(1) = "12345678901": Check Rejected(id, values), "Overlength DoD ID accepted"
    For i = 2 To 5
        values = Seed(): values(i) = "yes": Check Rejected(id, values), "Case variant component accepted"
        values(i) = True: Check Rejected(id, values), "Numeric/Boolean component code accepted"
    Next
    For Each bad In Array("2023-02-29", "2024-02-30", "2026-13-01", "2026-00-01", "2026-10-00", "10/03/2026", "2026-1-01", "2026-10-03 ", "2026-10-03T12:00:00", 20261003)
        values = Seed(): values(6) = bad: Check Rejected(id, values), "Invalid/ambiguous date accepted"
    Next
    Check Rejected(id, Array(Null)), "Malformed member array accepted"
    Check DCount("*", "ReportRevision") = before, "Invalid input wrote history"
    values = Array("001", "", Null, "", "No", "Yes", Null): Call SaveMember(id, "FITREP", values)
    Check Stored(id, "MemberSSN") = "001" And Not IsNull(Stored(id, "MemberDoDID")) And IsNull(Stored(id, "DateReported")), "Incomplete/empty/null values changed"
    before = DCount("*", "ReportRevision")
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<> 'DateReported'"
    Check Rejected(id, Seed()), "Late write failure accepted"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = ""
    Check Stored(id, "MemberSSN") = "001" And DCount("*", "ReportRevision") = before, "Late failure did not roll back features"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Call OpenMemberDetails()
    Forms!frmMemberDetails!txtMemberSSN = "PENDING INVALID IDENTIFIER"
    Forms!frmMemberDetails!txtComponentActive = "PENDING INVALID COMPONENT"
    Forms!frmMemberDetails!txtDateReported = "PENDING INVALID DATE"
    Call ApplyMemberDetails(): Call FixtureCloseEditor()
    Check RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft")), "Recovery failed"
    Check Forms!frmEditor!txtMemberSSN = "PENDING INVALID IDENTIFIER" And Forms!frmEditor!txtComponentActive = "PENDING INVALID COMPONENT" And Forms!frmEditor!txtDateReported = "PENDING INVALID DATE", "Raw recovery truncated"
    Check Stored(id, "MemberSSN") = "001" And IsNull(Stored(id, "DateReported")), "Recovery auto committed"
    Call FixtureCloseEditor()
    Call ArchiveDraft(id, True): Check Rejected(id, Seed()), "Archive guard absent"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Call OpenMemberDetails()
    Check Not Forms!frmMemberDetails!cmdApply.Enabled, "Archived panel can apply"
    Call CancelMemberDetails(): Call FixtureCloseEditor(): Call ArchiveDraft(id, False)
    first = DMin("RevisionID", "ReportRevision"): copy = DLookup("ReportID", "ReportRevision", "RevisionID=" & first)
    For Each key In keys: CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & first & " AND FieldKey='" & key & "'", dbFailOnError: Next
    CurrentDb.Execute "UPDATE ReportRevision SET SnapshotSchema=14 WHERE RevisionID=" & first, dbFailOnError
    Check OpenRevisionComparison(copy, first), "Legacy comparison failed"
    Forms!frmRevisionCompare!chkChangedOnly = False: Call RefreshComparisonFields()
    Forms!frmRevisionCompare!lstCompareFields = "DateReported": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtOldValue = "[Not recorded in this revision]", "Legacy marker absent"
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo: Call RestoreDraft(copy, first)
    For Each key In keys: Check IsNull(Stored(copy, CStr(key))), "Legacy restore invented details": Next
    second = DLookup("CurrentRevisionID", "Report", "ReportID=" & copy)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & second & " AND FieldKey='DateReported'", dbFailOnError
    Check RestoreRejected(copy, second), "Corrupt new snapshot restored"
    Set r = CurrentDb.OpenRecordset("RevisionValue", dbOpenDynaset)
    r.AddNew: r!RevisionID = second: r!FieldKey = "DateReported": r!IsNullValue = True: r.Update: r.Close
    TestMemberDetails = "PASS: " & checks & " member details assertions": Exit Function
Failed:
    TestMemberDetails = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function VerifyMemberDetailsRecovery() As Boolean
    Call RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft"))
    VerifyMemberDetailsRecovery = (Forms!frmEditor!txtMemberSSN = "PENDING INVALID IDENTIFIER" And Forms!frmEditor!txtComponentActive = "PENDING INVALID COMPONENT" And Forms!frmEditor!txtDateReported = "PENDING INVALID DATE")
    Call FixtureCloseEditor()
End Function

Private Sub FixtureCloseEditor()
    ' Preserve the old abrupt-close semantics for persistence/recovery fixtures.
    ' Real save/discard/cancel exits are exercised by UXWorkflowHarness.
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Sub
