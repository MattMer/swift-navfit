Attribute VB_Name = "AssessmentHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 2250, , message
    checks = checks + 1
End Sub
Private Function Stored(ByVal id As Long, ByVal key As String) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT [" & key & "] FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Stored = r.Fields(0).Value: r.Close
End Function
Private Function SaveAssessment(ByVal id As Long, ByVal kind As String, ByVal values As Variant) As Long
    SaveAssessment = CommitDraft(id, kind, "Assessment", "Synthetic", Null, Null, "", "Assessment test", "Narrative", , , values)
End Function
Private Function Rejected(ByVal id As Long, ByVal values As Variant) As Boolean
    On Error GoTo Failed
    Call SaveAssessment(id, "FITREP", values): Exit Function
Failed:
    Rejected = True
End Function
Private Function RestoreRejected(ByVal id As Long, ByVal revision As Long) As Boolean
    On Error GoTo Failed
    Call RestoreDraft(id, revision): Exit Function
Failed:
    RestoreRejected = True
End Function
Private Function ReviewRejected(ByVal task As Long) As Boolean
    On Error GoTo Failed
    Call FinishReview(task, "Accepted", ""): Exit Function
Failed:
    ReviewRejected = True
End Function
Public Function TestAssessment() As String
    Dim kind As Variant, id As Long, first As Long, second As Long, copy As Long, task As Long, bad As Variant
    Dim before As Long, recovery As Long, r As DAO.Recordset, choice As Variant
    On Error GoTo Failed
    Call ConfigureWorkspace("Assessment tests", "Synthetic reviewer")
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        id = SaveAssessment(0, CStr(kind), Array("Observed", "Promotable"))
        first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check Stored(id, "ObservationStatus") = "Observed" And Stored(id, "PromotionRecommendation") = "Promotable", "Selections lost"
        Check DCount("*", "RevisionValue", "RevisionID=" & first) = 43, "Incomplete snapshot"
        Call CommitDraft(id, CStr(kind), "Assessment", "Synthetic", Null, Null, "", "Omitted", "Narrative")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Omitted assessment changed history"
        copy = DuplicateDraft(id)
        Check Stored(copy, "ObservationStatus") = "Observed" And Stored(copy, "PromotionRecommendation") = "Promotable", "Duplicate lost assessment"
        task = AssignReview(id, first, DLookup("ActiveProfileID", "Workspace"), Null, "Review")
        Call SaveAssessment(id, CStr(kind), Array("NOB", "NOB"))
        second = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check ReviewRejected(task), "Assessment edit left review current"
        Call FinishReview(task, "Cancelled", "New assessment")
        Check RevisionFieldStatus(first, second, "ObservationStatus") = "Changed", "Observation change invisible"
        Check RevisionFieldStatus(first, second, "PromotionRecommendation") = "Changed", "Recommendation change invisible"
        Check OpenRevisionComparison(id, first), "Comparison failed"
        Forms!frmRevisionCompare!lstCompareFields = "PromotionRecommendation": Call LoadComparisonValues()
        Check Forms!frmRevisionCompare!txtOldValue = "Promotable" And Forms!frmRevisionCompare!txtNewValue = "NOB", "Comparison values wrong"
        DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
        Call RestoreDraft(id, first)
        Check Stored(id, "PromotionRecommendation") = "Promotable", "Restore lost assessment"
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Check Forms!frmEditor!cboObservationStatus = "Observed", "Editor did not load assessment"
        Forms!frmEditor!cboPromotionRecommendation = "Early Promote"
        Check SaveEditor(), "Editor save failed"
        Check Stored(id, "PromotionRecommendation") = "Early Promote", "Editor save lost recommendation"
        Forms!frmEditor!cboTrait33 = "NOB": Call ShowTraitGuidance(33)
        Check Forms!frmEditor!cboObservationStatus = "Observed" And Forms!frmEditor!cboPromotionRecommendation = "Early Promote", "Traits inferred assessment"
        DoCmd.Close acForm, "frmEditor", acSaveNo
    Next
    For Each choice In Array("NOB", "Significant Problems", "Progressing", "Promotable", "Must Promote", "Early Promote")
        Call SaveAssessment(id, "FITREP", Array("Observed", CStr(choice)))
        Check Stored(id, "PromotionRecommendation") = choice, "Listed recommendation rejected"
    Next
    Call SaveAssessment(id, "FITREP", Array(Null, ""))
    Check IsNull(Stored(id, "ObservationStatus")) And Not IsNull(Stored(id, "PromotionRecommendation")) And Len(Stored(id, "PromotionRecommendation")) = 0, "Null/empty collapsed"
    before = DCount("*", "ReportRevision")
    For Each bad In Array("EP", "early promote", " Early Promote", "Unlisted", 6, vbTab)
        Check Rejected(id, Array("Observed", bad)), "Invalid recommendation accepted"
    Next
    Check Rejected(id, Array("Unknown", "Promotable")), "Unknown observation accepted"
    Check Rejected(id, Array("Observed")), "Incomplete assessment array accepted"
    Check DCount("*", "ReportRevision") = before, "Invalid input wrote revisions"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<> 'PromotionRecommendation'"
    Check Rejected(id, Array("NOB", "NOB")), "Late write failure accepted"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = ""
    Check IsNull(Stored(id, "ObservationStatus")) And DCount("*", "ReportRevision") = before, "Late failure did not roll back"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!cboObservationStatus = "UNSUPPORTED PENDING": Forms!frmEditor!cboPromotionRecommendation = "Must Promote"
    Check CheckpointEditor(), "Checkpoint failed"
    recovery = DMax("RecoveryDraftID", "RecoveryDraft")
    Check DLookup("ObservationStatus", "RecoveryDraft", "RecoveryDraftID=" & recovery) = "UNSUPPORTED PENDING", "Raw checkpoint truncated"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    Check RestoreRecoveryDraft(recovery), "Recovery failed"
    Check Forms!frmEditor!cboObservationStatus = "UNSUPPORTED PENDING" And Forms!frmEditor!cboPromotionRecommendation = "Must Promote", "Recovery changed values"
    Check IsNull(Stored(id, "ObservationStatus")), "Recovery auto committed"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    Call ArchiveDraft(id, True)
    Check Rejected(id, Array("Observed", "Promotable")), "Archive guard absent"
    Call ArchiveDraft(id, False)
    'Simulate schema 10 history, keeping its actual 19-field snapshots.
    first = DMin("RevisionID", "ReportRevision"): copy = DLookup("ReportID", "ReportRevision", "RevisionID=" & first)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & first & " AND FieldKey In ('ObservationStatus','PromotionRecommendation')", dbFailOnError
    CurrentDb.Execute "UPDATE ReportRevision SET SnapshotSchema=10 WHERE RevisionID=" & first, dbFailOnError
    Check OpenRevisionComparison(copy, first), "Legacy comparison failed"
    Forms!frmRevisionCompare!chkChangedOnly = False: Call RefreshComparisonFields()
    Forms!frmRevisionCompare!lstCompareFields = "ObservationStatus": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtOldValue = "[Not recorded in this revision]", "Legacy marker absent"
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
    Call RestoreDraft(copy, first)
    Check IsNull(Stored(copy, "ObservationStatus")) And IsNull(Stored(copy, "PromotionRecommendation")), "Legacy restore invented assessment"
    second = DLookup("CurrentRevisionID", "Report", "ReportID=" & copy)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & second & " AND FieldKey='ObservationStatus'", dbFailOnError
    Check RestoreRejected(copy, second), "Corrupt new snapshot restored"
    Set r = CurrentDb.OpenRecordset("RevisionValue", dbOpenDynaset)
    r.AddNew: r!RevisionID = second: r!FieldKey = "ObservationStatus": r!IsNullValue = True: r.Update: r.Close
    TestAssessment = "PASS: " & checks & " assessment assertions": Exit Function
Failed:
    TestAssessment = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function VerifyAssessmentRecovery() As Boolean
    Call RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft"))
    VerifyAssessmentRecovery = (Forms!frmEditor!cboObservationStatus = "UNSUPPORTED PENDING" And Forms!frmEditor!cboPromotionRecommendation = "Must Promote")
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Function
