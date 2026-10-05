Attribute VB_Name = "PromotionStatusHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 2350, , message
    checks = checks + 1
End Sub
Private Function Stored(ByVal id As Long) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT PromotionStatus FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Stored = r!PromotionStatus.Value: r.Close
End Function
Private Function SaveStatus(ByVal id As Long, ByVal kind As String, ByVal value As Variant) As Long
    SaveStatus = CommitDraft(id, kind, "Status", "Synthetic", Null, Null, "", "Status test", "Narrative", , , , value)
End Function
Private Function Rejected(ByVal id As Long, ByVal value As Variant) As Boolean
    On Error GoTo Failed
    Call SaveStatus(id, "FITREP", value): Exit Function
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
Public Function TestPromotionStatus() As String
    Dim kind As Variant, id As Long, first As Long, second As Long, copy As Long, task As Long, bad As Variant
    Dim before As Long, recovery As Long, r As DAO.Recordset, choice As Variant
    On Error GoTo Failed
    Call ConfigureWorkspace("Promotion status tests", "Synthetic reviewer")
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        id = SaveStatus(0, CStr(kind), "REGULAR")
        first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check Stored(id) = "REGULAR", "Status lost"
        Check DCount("*", "RevisionValue", "RevisionID=" & first) = 78, "Incomplete snapshot"
        Call CommitDraft(id, CStr(kind), "Status", "Synthetic", Null, Null, "", "Omitted", "Narrative")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1 And Stored(id) = "REGULAR", "Omitted status changed report"
        copy = DuplicateDraft(id): Check Stored(copy) = "REGULAR", "Duplicate lost status"
        task = AssignReview(id, first, DLookup("ActiveProfileID", "Workspace"), Null, "Review")
        Call SaveStatus(id, CStr(kind), "FROCKED")
        second = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check ReviewRejected(task), "Status edit left review current"
        Call FinishReview(task, "Cancelled", "New status")
        Check RevisionFieldStatus(first, second, "PromotionStatus") = "Changed", "Status change invisible"
        Check OpenRevisionComparison(id, first), "Comparison failed"
        Forms!frmRevisionCompare!lstCompareFields = "PromotionStatus": Call LoadComparisonValues()
        Check Forms!frmRevisionCompare!txtOldValue = "REGULAR" And Forms!frmRevisionCompare!txtNewValue = "FROCKED", "Comparison values wrong"
        DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
        Call RestoreDraft(id, first): Check Stored(id) = "REGULAR", "Restore lost status"
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Check Forms!frmEditor!cboPromotionStatus = "REGULAR", "Editor did not load status"
        Forms!frmEditor!cboPromotionStatus = "SELECTED"
        Check SaveEditor(), "Editor save failed"
        Check Stored(id) = "SELECTED", "Editor save lost status"
        Forms!frmEditor!cboPromotionRecommendation = "Early Promote": Forms!frmEditor!cboTrait33 = "NOB"
        Call ShowTraitGuidance(33)
        Check Forms!frmEditor!cboPromotionStatus = "SELECTED", "Recommendation/traits inferred status"
        Call FixtureCloseEditor()
    Next
    For Each choice In Array("REGULAR", "FROCKED", "SELECTED", "SPOT")
        Call SaveStatus(id, "FITREP", choice): Check Stored(id) = choice, "Listed status rejected"
    Next
    Call SaveStatus(id, "FITREP", ""): Check Not IsNull(Stored(id)) And Len(Stored(id)) = 0, "Empty status collapsed"
    Call SaveStatus(id, "FITREP", Null): Check IsNull(Stored(id)), "Null status changed"
    before = DCount("*", "ReportRevision")
    For Each bad In Array("regular", " REGULAR", "REGULAR ", "UNKNOWN", 2, True, vbTab, "FROCKED" & vbCrLf)
        Check Rejected(id, bad), "Unsupported status accepted"
    Next
    Check DCount("*", "ReportRevision") = before, "Invalid input wrote history"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<> 'PromotionStatus'"
    Check Rejected(id, "SPOT"), "Late write failure accepted"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = ""
    Check IsNull(Stored(id)) And DCount("*", "ReportRevision") = before, "Late failure did not roll back"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!cboPromotionStatus = "UNSUPPORTED PENDING"
    Check CheckpointEditor(), "Checkpoint failed"
    recovery = DMax("RecoveryDraftID", "RecoveryDraft")
    Check DLookup("PromotionStatus", "RecoveryDraft", "RecoveryDraftID=" & recovery) = "UNSUPPORTED PENDING", "Raw checkpoint truncated"
    Call FixtureCloseEditor()
    Check RestoreRecoveryDraft(recovery), "Recovery failed"
    Check Forms!frmEditor!cboPromotionStatus = "UNSUPPORTED PENDING", "Recovery changed status"
    Check IsNull(Stored(id)), "Recovery auto committed"
    Call FixtureCloseEditor()
    Call ArchiveDraft(id, True): Check Rejected(id, "REGULAR"), "Archived status edit accepted"
    Call ArchiveDraft(id, False)
    first = DMin("RevisionID", "ReportRevision"): copy = DLookup("ReportID", "ReportRevision", "RevisionID=" & first)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & first & " AND FieldKey='PromotionStatus'", dbFailOnError
    CurrentDb.Execute "UPDATE ReportRevision SET SnapshotSchema=11 WHERE RevisionID=" & first, dbFailOnError
    Check OpenRevisionComparison(copy, first), "Legacy comparison failed"
    Forms!frmRevisionCompare!chkChangedOnly = False: Call RefreshComparisonFields()
    Forms!frmRevisionCompare!lstCompareFields = "PromotionStatus": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtOldValue = "[Not recorded in this revision]", "Legacy marker absent"
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
    Call RestoreDraft(copy, first): Check IsNull(Stored(copy)), "Legacy restore invented status"
    second = DLookup("CurrentRevisionID", "Report", "ReportID=" & copy)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & second & " AND FieldKey='PromotionStatus'", dbFailOnError
    Check RestoreRejected(copy, second), "Corrupt new snapshot restored"
    Set r = CurrentDb.OpenRecordset("RevisionValue", dbOpenDynaset)
    r.AddNew: r!RevisionID = second: r!FieldKey = "PromotionStatus": r!IsNullValue = True: r.Update: r.Close
    TestPromotionStatus = "PASS: " & checks & " promotion status assertions": Exit Function
Failed:
    TestPromotionStatus = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function VerifyPromotionStatusRecovery() As Boolean
    Call RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft"))
    VerifyPromotionStatusRecovery = (Forms!frmEditor!cboPromotionStatus = "UNSUPPORTED PENDING")
    Call FixtureCloseEditor()
End Function

Private Sub FixtureCloseEditor()
    ' Preserve the old abrupt-close semantics for persistence/recovery fixtures.
    ' Real save/discard/cancel exits are exercised by UXWorkflowHarness.
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Sub
