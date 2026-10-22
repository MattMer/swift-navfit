Attribute VB_Name = "TraitsHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 2050, , message
    checks = checks + 1
    CurrentDb.Execute "UPDATE AppMetadata SET CompatibilityStatus='TEST: " & Replace(message, "'", "''") & "' WHERE MetadataID=1", dbFailOnError
End Sub
Private Function Stored(ByVal id As Long, ByVal key As String) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT [" & key & "] FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Stored = r.Fields(0).Value: r.Close
End Function
Private Function SaveTraits(ByVal id As Long, ByVal kind As String, ByVal values As Variant) As Long
    SaveTraits = CommitDraft(id, kind, "Synthetic", "Synthetic Member", Null, Null, "Notes", "Traits", "Narrative", , values)
End Function
Private Function Rejected(ByVal id As Long, ByVal kind As String, ByVal values As Variant) As Boolean
    On Error GoTo Failed
    Call SaveTraits(id, kind, values): Exit Function
Failed:
    Rejected = True
End Function
Private Function StaleRejected(ByVal task As Long) As Boolean
    On Error GoTo Failed
    Call FinishReview(task, "Accepted", "Stale"): Exit Function
Failed:
    StaleRejected = InStr(Err.Description, "stale") > 0
End Function
Private Function CorruptRejected(ByVal id As Long, ByVal revision As Long) As Boolean
    On Error GoTo Failed
    Call RestoreDraft(id, revision): Exit Function
Failed:
    CorruptRejected = InStr(Err.Description, "incomplete") > 0
End Function
Public Function TestTraits() As String
    Dim kind As Variant, id As Long, first As Long, second As Long, copy As Long, values As Variant, bad As Variant
    Dim i As Long, before As Long, task As Long, recovery As Long, r As DAO.Recordset
    On Error GoTo Failed
    If DCount("*", "Workspace") > 0 Then Err.Raise vbObjectError + 2051, , "Trait fixture requires a blank test copy."
    Call ConfigureWorkspace("SYNTHETIC TRAITS", "Synthetic Reviewer")
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        values = Array("NOB", "1.0", "2.0", "3.0", "4.0", "5.0", Null)
        id = SaveTraits(0, CStr(kind), values)
        first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check Stored(id, "Trait33") = "NOB" And IsNull(Stored(id, "Trait39")), "NOB and blank collapsed"
        Check DCount("*", "RevisionValue", "RevisionID=" & first) = SavedFields(id).Count, "Trait snapshot incomplete"
        Check DCount("*", "RevisionValue", "RevisionID=" & first & " AND FieldKey In ('Trait33','Trait34','Trait35','Trait36','Trait37','Trait38','Trait39')") = 7, "Seven trait snapshot keys missing"
        Call CommitDraft(id, CStr(kind), "Synthetic", "Synthetic Member", Null, Null, "Notes", "Omitted", "Narrative")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1 And Stored(id, "Trait34") = "1.0", "Omitted traits changed report"
        copy = DuplicateDraft(id)
        Check Stored(copy, "Trait33") = "NOB" And Stored(copy, "Trait38") = "5.0", "Duplicate lost traits"
        task = AssignReview(id, first, DLookup("ActiveProfileID", "Workspace"), Null, "Trait review")
        values(1) = "2.0": Call SaveTraits(id, CStr(kind), values)
        second = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check RevisionFieldStatus(first, second, "Trait34") = "Changed", "Trait change invisible"
        Check StaleRejected(task), "Trait save did not stale review"
        Call FinishReview(task, "Cancelled", "New marks")
        Check OpenRevisionComparison(id, first), "Trait comparison failed"
        Forms!frmRevisionCompare!lstCompareFields = "Trait34": Call LoadComparisonValues()
        Check Forms!frmRevisionCompare!txtOldValue = "1.0" And Forms!frmRevisionCompare!txtNewValue = "2.0", "Trait comparison values wrong"
        DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
        Call RestoreDraft(id, first): Check Stored(id, "Trait34") = "1.0", "Restore lost traits"
        values(6) = "": Call SaveTraits(id, CStr(kind), values)
        Check Not IsNull(Stored(id, "Trait39")) And Len(Stored(id, "Trait39")) = 0, "Empty trait changed to null"
        ' Current fixed-page/guidance UI is covered by CompatPages/ReportTools.
    Next
    before = DCount("*", "ReportRevision")
    For Each bad In Array("4.5", "0.0", "nob", "5", "6.0", vbTab, 4)
        values = Array(Null, Null, Null, Null, Null, Null, Null): values(0) = bad
        Check Rejected(id, "FITREP", values), "Unsupported mark accepted"
    Next
    Check Rejected(id, "FITREP", Array("NOB")), "Incomplete trait array accepted"
    Check DCount("*", "ReportRevision") = before, "Rejected marks changed history"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<>'Trait39'"
    values = Array("5.0", "5.0", "5.0", "5.0", "5.0", "5.0", "5.0")
    Check Rejected(id, "FITREP", values), "Late trait snapshot failure accepted"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = ""
    Check DCount("*", "ReportRevision") = before And Stored(id, "Trait33") = "NOB", "Failed trait save partly committed"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!cboTrait33 = "PENDING INVALID": Forms!frmEditor!cboTrait39 = "NOB"
    Check CheckpointEditor(), "Trait checkpoint failed"
    recovery = DMax("RecoveryDraftID", "RecoveryDraft")
    Check DLookup("Trait33", "RecoveryDraft", "RecoveryDraftID=" & recovery) = "PENDING INVALID", "Pending mark truncated"
    Call FixtureCloseEditor()
    Check RestoreRecoveryDraft(recovery), "Trait recovery failed"
    Check Forms!frmEditor!cboTrait33 = "PENDING INVALID" And Forms!frmEditor!cboTrait39 = "NOB", "Recovered marks changed"
    Check Stored(id, "Trait33") = "NOB", "Recovery auto-saved invalid mark"
    Call FixtureCloseEditor()
    Call ArchiveDraft(id, True): Check Rejected(id, "FITREP", values), "Archived trait edit accepted"
    Call ArchiveDraft(id, False)
    'Simulate an actual schema-9 revision, preserving its other 12 snapshot fields.
    first = DMin("RevisionID", "ReportRevision"): copy = DLookup("ReportID", "ReportRevision", "RevisionID=" & first)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & first & " AND FieldKey In ('Trait33','Trait34','Trait35','Trait36','Trait37','Trait38','Trait39')", dbFailOnError
    CurrentDb.Execute "UPDATE ReportRevision SET SnapshotSchema=9 WHERE RevisionID=" & first, dbFailOnError
    Check OpenRevisionComparison(copy, first), "Legacy trait comparison failed"
    Forms!frmRevisionCompare!chkChangedOnly = False: Call RefreshComparisonFields()
    Forms!frmRevisionCompare!lstCompareFields = "Trait33": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtOldValue = "[Not recorded in this revision]", "Legacy trait marker absent"
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
    Call RestoreDraft(copy, first): Check IsNull(Stored(copy, "Trait33")), "Legacy restore invented marks"
    second = DLookup("CurrentRevisionID", "Report", "ReportID=" & copy)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & second & " AND FieldKey='Trait33'", dbFailOnError
    Check CorruptRejected(copy, second), "Incomplete new trait snapshot restored"
    Set r = CurrentDb.OpenRecordset("RevisionValue", dbOpenDynaset)
    r.AddNew: r!RevisionID = second: r!FieldKey = "Trait33": r!IsNullValue = True: r.Update: r.Close
    TestTraits = "PASS: " & checks & " trait assertions": Exit Function
Failed:
    TestTraits = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function VerifyTraitRecovery() As Boolean
    Dim recovery As Long
    recovery = DMax("RecoveryDraftID", "RecoveryDraft")
    Call RestoreRecoveryDraft(recovery)
    VerifyTraitRecovery = (Forms!frmEditor!cboTrait33 = "PENDING INVALID" And Forms!frmEditor!cboTrait39 = "NOB")
    Call FixtureCloseEditor()
End Function

Private Sub FixtureCloseEditor()
    ' Preserve the old abrupt-close semantics for persistence/recovery fixtures.
    ' Real save/discard/cancel exits are exercised by UXWorkflowHarness.
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Sub
