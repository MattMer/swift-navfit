Attribute VB_Name = "ServiceDetailsHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 2750, , message
    checks = checks + 1
End Sub
Private Function Seed() As Variant
    Seed = Array("2024-02-29", "O'NEIL, A", "P", "Good", "2026-01-01", "BASIC")
End Function
Private Function Stored(ByVal id As Long, ByVal key As String) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT [" & key & "] FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Stored = r.Fields(0).Value: r.Close
End Function
Private Function SaveService(ByVal id As Long, ByVal kind As String, ByVal values As Variant) As Long
    SaveService = CommitDraft(id, kind, "Service", "Synthetic", Null, Null, "", "Service test", "Narrative", , , , , , , , values)
End Function
Private Function Rejected(ByVal id As Long, ByVal values As Variant) As Boolean
    On Error GoTo Failed
    Call SaveService(id, "FITREP", values): Exit Function
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
Public Function TestServiceDetails() As String
    Dim kind As Variant, values As Variant, keys As Variant, key As Variant, bad As Variant, index As Variant
    Dim id As Long, first As Long, second As Long, copy As Long, task As Long, before As Long, i As Long, r As DAO.Recordset
    Dim limits As Variant
    On Error GoTo Failed
    Call ConfigureWorkspace("Service tests", "Synthetic reviewer"): keys = ServiceDetailsKeys(): limits = Array(10, 20, 4, 15, 10, 10)
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        values = Seed(): id = SaveService(0, CStr(kind), values): first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check DCount("*", "RevisionValue", "RevisionID=" & first) = 78, "Incomplete snapshot"
        For i = 0 To 5: Check Stored(id, CStr(keys(i))) = values(i), "Saved feature changed": Next
        Call CommitDraft(id, CStr(kind), "Service", "Synthetic", Null, Null, "", "Omitted", "Narrative")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Omitted details wrote history"
        copy = DuplicateDraft(id)
        Check Stored(copy, "CounselorName") = values(1) And Stored(copy, "ReadinessDetail") = "Good" And Stored(copy, "BilletSubcategory") = "BASIC", "Duplicate lost features"
        For Each index In Array(0, 2, 5)
            first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            task = AssignReview(id, first, DLookup("ActiveProfileID", "Workspace"), Null, "Feature review")
            Select Case index
                Case 0: values(0) = "2026-10-03"
                Case 2: values(2) = "F"
                Case 5: values(5) = "OIC"
            End Select
            Call SaveService(id, CStr(kind), values): second = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            Check ReviewRejected(task), "Feature edit left review current"
            Call FinishReview(task, "Cancelled", "Changed details")
            Check RevisionFieldStatus(first, second, CStr(keys(index))) = "Changed", "Feature change invisible"
            Check OpenRevisionComparison(id, first), "Comparison failed"
            Forms!frmRevisionCompare!lstCompareFields = keys(index): Call LoadComparisonValues()
            Check Forms!frmRevisionCompare!txtNewValue = values(index), "Comparison value wrong"
            DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
        Next
        first = DMin("RevisionID", "ReportRevision", "ReportID=" & id): Call RestoreDraft(id, first)
        Check Stored(id, "CounselingDate") = "2024-02-29" And Stored(id, "ReadinessSummary") = "P" And Stored(id, "BilletSubcategory") = "BASIC", "Restore lost features"
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Check OpenServiceDetails(), "Panel failed"
        Check Forms!frmServiceDetails!txtCounselorName = "O'NEIL, A", "Panel lost counselor"
        Forms!frmServiceDetails!txtCounselorName = "CANCEL ME": Call CancelServiceDetails()
        Check Forms!frmEditor!txtCounselorName = "O'NEIL, A", "Cancel changed pending values"
        Call OpenServiceDetails(): Forms!frmServiceDetails!txtCounselorName = "NEW COUNSELOR"
        Forms!frmServiceDetails!txtReadinessDetail = "Excellent": Forms!frmServiceDetails!txtBilletSubcategory = "OIC"
        Check ApplyServiceDetails(), "Apply failed"
        Check Stored(id, "CounselorName") = "O'NEIL, A" And Forms!frmEditor!txtCounselorName = "NEW COUNSELOR", "Apply committed revision"
        Check DLookup("ReadinessDetail", "RecoveryDraft", "ReportID=" & id) = "Excellent", "Apply did not checkpoint"
        Check SaveEditor(), "Editor save failed"
        Check Stored(id, "CounselorName") = "NEW COUNSELOR" And Stored(id, "ReadinessDetail") = "Excellent" And Stored(id, "BilletSubcategory") = "OIC", "Editor save lost features"
        DoCmd.Close acForm, "frmEditor", acSaveNo
    Next
    For Each index In Array(1, 2, 3, 5)
        values = Seed(): values(index) = String$(limits(index), "X")
        Call SaveService(id, "FITREP", values): Check Stored(id, CStr(keys(index))) = values(index), "Boundary truncated"
        values(index) = values(index) & "X": Check Rejected(id, values), "Overlength accepted"
    Next
    before = DCount("*", "ReportRevision")
    For i = 0 To 5
        values = Seed(): values(i) = 123: Check Rejected(id, values), "Nontext accepted"
        values(i) = "X" & vbTab: Check Rejected(id, values), "Tab accepted"
        values(i) = "X" & vbCrLf: Check Rejected(id, values), "Multiline accepted"
    Next
    For Each index In Array(0, 4)
        For Each bad In Array("2023-02-29", "1900-02-29", "2024-02-30", "2026-13-01", "2026-10-00", "10/03/2026", "2026-10-03T12:00", "2026-10-03 ")
            values = Seed(): values(index) = bad: Check Rejected(id, values), "Invalid date accepted"
        Next
    Next
    Check Rejected(id, Array(Null)), "Malformed array accepted"
    Check DCount("*", "ReportRevision") = before, "Rejected save wrote history"
    values = Array("2000-02-29", "", Null, "", "2000-02-29", Null): Call SaveService(id, "FITREP", values)
    Check Stored(id, "CounselingDate") = "2000-02-29" And Stored(id, "ReadinessDate") = "2000-02-29", "Century leap date rejected"
    Check IsNull(Stored(id, "ReadinessSummary")) And Not IsNull(Stored(id, "ReadinessDetail")), "Null/empty collapsed"
    before = DCount("*", "ReportRevision")
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<> 'BilletSubcategory'"
    Check Rejected(id, Seed()), "Late write failure accepted"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = ""
    Check IsNull(Stored(id, "ReadinessSummary")) And DCount("*", "ReportRevision") = before, "Late failure did not roll back features"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Call OpenServiceDetails()
    Forms!frmServiceDetails!txtCounselingDate = "PENDING INVALID DATE"
    Forms!frmServiceDetails!txtReadinessDetail = "PENDING OVERLONG READINESS"
    Forms!frmServiceDetails!txtBilletSubcategory = "PENDING OVERLONG BILLET"
    Call ApplyServiceDetails(): DoCmd.Close acForm, "frmEditor", acSaveNo
    Check RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft")), "Recovery failed"
    Check Forms!frmEditor!txtCounselingDate = "PENDING INVALID DATE" And Forms!frmEditor!txtReadinessDetail = "PENDING OVERLONG READINESS" And Forms!frmEditor!txtBilletSubcategory = "PENDING OVERLONG BILLET", "Raw recovery truncated"
    Check Stored(id, "CounselingDate") = "2000-02-29" And IsNull(Stored(id, "BilletSubcategory")), "Recovery auto committed"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    Call ArchiveDraft(id, True): Check Rejected(id, Seed()), "Archive guard absent"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Call OpenServiceDetails()
    Check Not Forms!frmServiceDetails!cmdApply.Enabled, "Archived panel can apply"
    Call CancelServiceDetails(): DoCmd.Close acForm, "frmEditor", acSaveNo: Call ArchiveDraft(id, False)
    first = DMin("RevisionID", "ReportRevision"): copy = DLookup("ReportID", "ReportRevision", "RevisionID=" & first)
    For Each key In keys: CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & first & " AND FieldKey='" & key & "'", dbFailOnError: Next
    CurrentDb.Execute "UPDATE ReportRevision SET SnapshotSchema=15 WHERE RevisionID=" & first, dbFailOnError
    Check OpenRevisionComparison(copy, first), "Legacy comparison failed"
    Forms!frmRevisionCompare!chkChangedOnly = False: Call RefreshComparisonFields()
    Forms!frmRevisionCompare!lstCompareFields = "BilletSubcategory": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtOldValue = "[Not recorded in this revision]", "Legacy marker absent"
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo: Call RestoreDraft(copy, first)
    For Each key In keys: Check IsNull(Stored(copy, CStr(key))), "Legacy restore invented details": Next
    second = DLookup("CurrentRevisionID", "Report", "ReportID=" & copy)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & second & " AND FieldKey='BilletSubcategory'", dbFailOnError
    Check RestoreRejected(copy, second), "Corrupt new snapshot restored"
    Set r = CurrentDb.OpenRecordset("RevisionValue", dbOpenDynaset)
    r.AddNew: r!RevisionID = second: r!FieldKey = "BilletSubcategory": r!IsNullValue = True: r.Update: r.Close
    TestServiceDetails = "PASS: " & checks & " service details assertions": Exit Function
Failed:
    TestServiceDetails = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function VerifyServiceDetailsRecovery() As Boolean
    Call RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft"))
    VerifyServiceDetailsRecovery = (Forms!frmEditor!txtCounselingDate = "PENDING INVALID DATE" And Forms!frmEditor!txtReadinessDetail = "PENDING OVERLONG READINESS" And Forms!frmEditor!txtBilletSubcategory = "PENDING OVERLONG BILLET")
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Function
