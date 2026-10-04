Attribute VB_Name = "DecisionDetailsHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 2850, , message
    checks = checks + 1
End Sub
Private Function Seed() As Variant
    Seed = Array("DUTY A", "DUTY B", "Yes", "No", "Yes", "No")
End Function
Private Function Stored(ByVal id As Long, ByVal key As String) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT [" & key & "] FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Stored = r.Fields(0).Value: r.Close
End Function
Private Function SaveDecision(ByVal id As Long, ByVal kind As String, ByVal values As Variant) As Long
    SaveDecision = CommitDraft(id, kind, "Decision", "Synthetic", Null, Null, "", "Decision test", "Narrative", , , , , , , , , values)
End Function
Private Function Rejected(ByVal id As Long, ByVal values As Variant) As Boolean
    On Error GoTo Failed
    Call SaveDecision(id, "FITREP", values): Exit Function
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
Public Function TestDecisionDetails() As String
    Dim kind As Variant, values As Variant, keys As Variant, key As Variant, bad As Variant, index As Variant
    Dim id As Long, first As Long, second As Long, copy As Long, task As Long, before As Long, i As Long, r As DAO.Recordset
    Dim limits As Variant
    On Error GoTo Failed
    Call ConfigureWorkspace("Decision tests", "Synthetic reviewer"): keys = DecisionDetailsKeys(): limits = Array(20, 20, 3, 3, 3, 3)
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        values = Seed(): id = SaveDecision(0, CStr(kind), values): first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check DCount("*", "RevisionValue", "RevisionID=" & first) = 55, "Incomplete snapshot"
        For i = 0 To 5: Check Stored(id, CStr(keys(i))) = values(i), "Saved feature changed": Next
        Call CommitDraft(id, CStr(kind), "Decision", "Synthetic", Null, Null, "", "Omitted", "Narrative")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Omitted details wrote history"
        copy = DuplicateDraft(id)
        Check Stored(copy, "AssignmentRecommendationB") = values(1) And Stored(copy, "RetentionNo") = "No" And Stored(copy, "StatementNo") = "No", "Duplicate lost features"
        For Each index In Array(0, 2, 4)
            first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            task = AssignReview(id, first, DLookup("ActiveProfileID", "Workspace"), Null, "Feature review")
            Select Case index
                Case 0: values(0) = "NEW DUTY"
                Case 2: values(2) = "No"
                Case 4: values(4) = "No"
            End Select
            Call SaveDecision(id, CStr(kind), values): second = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            Check ReviewRejected(task), "Feature edit left review current"
            Call FinishReview(task, "Cancelled", "Changed details")
            Check RevisionFieldStatus(first, second, CStr(keys(index))) = "Changed", "Feature change invisible"
            Check OpenRevisionComparison(id, first), "Comparison failed"
            Forms!frmRevisionCompare!lstCompareFields = keys(index): Call LoadComparisonValues()
            Check Forms!frmRevisionCompare!txtNewValue = values(index), "Comparison value wrong"
            DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
        Next
        first = DMin("RevisionID", "ReportRevision", "ReportID=" & id): Call RestoreDraft(id, first)
        Check Stored(id, "AssignmentRecommendationA") = "DUTY A" And Stored(id, "RetentionYes") = "Yes" And Stored(id, "StatementNo") = "No", "Restore lost features"
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Check OpenDecisionDetails(), "Panel failed"
        Check Forms!frmDecisionDetails!txtAssignmentRecommendationB = "DUTY B", "Panel lost counselor"
        Forms!frmDecisionDetails!txtAssignmentRecommendationB = "CANCEL ME": Call CancelDecisionDetails()
        Check Forms!frmEditor!txtAssignmentRecommendationB = "DUTY B", "Cancel changed pending values"
        Call OpenDecisionDetails(): Forms!frmDecisionDetails!txtAssignmentRecommendationB = "NEW DUTY B"
        Forms!frmDecisionDetails!txtRetentionNo = "Yes": Forms!frmDecisionDetails!txtStatementNo = "Yes"
        Check ApplyDecisionDetails(), "Apply failed"
        Check Stored(id, "AssignmentRecommendationB") = "DUTY B" And Forms!frmEditor!txtAssignmentRecommendationB = "NEW DUTY B", "Apply committed revision"
        Check DLookup("RetentionNo", "RecoveryDraft", "ReportID=" & id) = "Yes", "Apply did not checkpoint"
        Check SaveEditor(), "Editor save failed"
        Check Stored(id, "AssignmentRecommendationB") = "NEW DUTY B" And Stored(id, "RetentionNo") = "Yes" And Stored(id, "StatementNo") = "Yes", "Editor save lost features"
        DoCmd.Close acForm, "frmEditor", acSaveNo
    Next
    For Each index In Array(0, 1)
        values = Seed(): values(index) = String$(limits(index), "X")
        Call SaveDecision(id, "FITREP", values): Check Stored(id, CStr(keys(index))) = values(index), "Boundary truncated"
        values(index) = values(index) & "X": Check Rejected(id, values), "Overlength accepted"
    Next
    For Each index In Array(2, 3, 4, 5)
        For Each bad In Array("Yes", "No", "", Null)
            values = Seed(): values(index) = bad: Call SaveDecision(id, "FITREP", values)
            If IsNull(bad) Then
                Check IsNull(Stored(id, CStr(keys(index)))), "Null flag lost"
            Else
                Check Not IsNull(Stored(id, CStr(keys(index)))), "Explicit entry became null"
                Check Stored(id, CStr(keys(index))) = bad, "Explicit entry changed"
            End If
        Next
    Next
    values = Array("A", "B", "Yes", "Yes", "Yes", "Yes"): Call SaveDecision(id, "FITREP", values)
    Check Stored(id, "RetentionYes") = "Yes" And Stored(id, "RetentionNo") = "Yes", "Paired retention flags inferred"
    Check Stored(id, "StatementYes") = "Yes" And Stored(id, "StatementNo") = "Yes", "Paired statement flags inferred"
    Check IsNull(Stored(id, "PromotionRecommendation")), "Assignment recommendation changed promotion recommendation"
    Call CommitDraft(id, "FITREP", "Decision", "Synthetic", Null, Null, "", "Separate promotion", "Narrative", , , Array("Observed", "Early Promote"))
    Check Stored(id, "AssignmentRecommendationA") = "A" And Stored(id, "RetentionYes") = "Yes", "Assessment save changed decisions"
    Check Stored(id, "PromotionRecommendation") = "Early Promote", "Promotion setup failed"
    values = Seed(): Call SaveDecision(id, "FITREP", values)
    Check Stored(id, "PromotionRecommendation") = "Early Promote", "Decision save changed promotion recommendation"
    before = DCount("*", "ReportRevision")
    For i = 0 To 5
        values = Seed(): values(i) = 123: Check Rejected(id, values), "Nontext accepted"
        values(i) = "X" & vbTab: Check Rejected(id, values), "Tab accepted"
        values(i) = "X" & vbCrLf: Check Rejected(id, values), "Multiline accepted"
    Next
    For Each index In Array(2, 3, 4, 5)
        For Each bad In Array("yes", "YES", "Y", "True", "0", "Yes ", "No ", "Bad")
            values = Seed(): values(index) = bad: Check Rejected(id, values), "Unsupported flag accepted"
        Next
    Next
    Check Rejected(id, Array(Null)), "Malformed array accepted"
    Check DCount("*", "ReportRevision") = before, "Rejected save wrote history"
    values = Array("", Null, Null, "", "No", Null): Call SaveDecision(id, "FITREP", values)
    Check Not IsNull(Stored(id, "AssignmentRecommendationA")) And IsNull(Stored(id, "AssignmentRecommendationB")), "Recommendation null/empty collapsed"
    Check IsNull(Stored(id, "RetentionYes")) And Not IsNull(Stored(id, "RetentionNo")), "Null/empty collapsed"
    before = DCount("*", "ReportRevision")
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<> 'StatementNo'"
    Check Rejected(id, Seed()), "Late write failure accepted"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = ""
    Check IsNull(Stored(id, "RetentionYes")) And DCount("*", "ReportRevision") = before, "Late failure did not roll back features"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Call OpenDecisionDetails()
    Forms!frmDecisionDetails!txtAssignmentRecommendationA = "PENDING OVERLONG RECOMMENDATION"
    Forms!frmDecisionDetails!txtRetentionNo = "UNKNOWN RETENTION"
    Forms!frmDecisionDetails!txtStatementNo = "UNKNOWN STATEMENT"
    Call ApplyDecisionDetails(): DoCmd.Close acForm, "frmEditor", acSaveNo
    Check RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft")), "Recovery failed"
    Check Forms!frmEditor!txtAssignmentRecommendationA = "PENDING OVERLONG RECOMMENDATION" And Forms!frmEditor!txtRetentionNo = "UNKNOWN RETENTION" And Forms!frmEditor!txtStatementNo = "UNKNOWN STATEMENT", "Raw recovery truncated"
    Check Stored(id, "AssignmentRecommendationA") = "" And IsNull(Stored(id, "StatementNo")), "Recovery auto committed"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    Call ArchiveDraft(id, True): Check Rejected(id, Seed()), "Archive guard absent"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Call OpenDecisionDetails()
    Check Not Forms!frmDecisionDetails!cmdApply.Enabled, "Archived panel can apply"
    Call CancelDecisionDetails(): DoCmd.Close acForm, "frmEditor", acSaveNo: Call ArchiveDraft(id, False)
    first = DMin("RevisionID", "ReportRevision"): copy = DLookup("ReportID", "ReportRevision", "RevisionID=" & first)
    For Each key In keys: CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & first & " AND FieldKey='" & key & "'", dbFailOnError: Next
    CurrentDb.Execute "UPDATE ReportRevision SET SnapshotSchema=16 WHERE RevisionID=" & first, dbFailOnError
    Check OpenRevisionComparison(copy, first), "Legacy comparison failed"
    Forms!frmRevisionCompare!chkChangedOnly = False: Call RefreshComparisonFields()
    Forms!frmRevisionCompare!lstCompareFields = "StatementNo": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtOldValue = "[Not recorded in this revision]", "Legacy marker absent"
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo: Call RestoreDraft(copy, first)
    For Each key In keys: Check IsNull(Stored(copy, CStr(key))), "Legacy restore invented details": Next
    second = DLookup("CurrentRevisionID", "Report", "ReportID=" & copy)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & second & " AND FieldKey='StatementNo'", dbFailOnError
    Check RestoreRejected(copy, second), "Corrupt new snapshot restored"
    Set r = CurrentDb.OpenRecordset("RevisionValue", dbOpenDynaset)
    r.AddNew: r!RevisionID = second: r!FieldKey = "StatementNo": r!IsNullValue = True: r.Update: r.Close
    TestDecisionDetails = "PASS: " & checks & " decision details assertions": Exit Function
Failed:
    TestDecisionDetails = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function VerifyDecisionDetailsRecovery() As Boolean
    Call RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft"))
    VerifyDecisionDetailsRecovery = (Forms!frmEditor!txtAssignmentRecommendationA = "PENDING OVERLONG RECOMMENDATION" And Forms!frmEditor!txtRetentionNo = "UNKNOWN RETENTION" And Forms!frmEditor!txtStatementNo = "UNKNOWN STATEMENT")
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Function
