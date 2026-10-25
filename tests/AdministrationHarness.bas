Attribute VB_Name = "AdministrationHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 1950, , message
    checks = checks + 1
    CurrentDb.Execute "UPDATE AppMetadata SET CompatibilityStatus='TEST: " & Replace(message, "'", "''") & "' WHERE MetadataID=1", dbFailOnError
End Sub
Private Function Stored(ByVal id As Long, ByVal key As String) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT [" & key & "] FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Stored = r.Fields(0).Value: r.Close
End Function
Private Function MissingInput(ByVal id As Long, ByVal key As String) As Boolean
    Dim item As Object
    For Each item In ValidateLiveFields(SavedFields(id), id)
        If item("Key") = key And item("Severity") = "Incomplete draft" Then MissingInput = True
    Next
End Function
Private Function Rejected(ByVal id As Long, ByVal kind As String, ByVal values As Variant) As Boolean
    On Error GoTo Failed
    Call CommitDraft(id, kind, "Synthetic", "Local display", Null, Null, "Notes", "Rejected", "Narrative", values)
    Exit Function
Failed:
    Rejected = True
End Function
Private Function StaleReviewRejected(ByVal id As Long) As Boolean
    On Error GoTo Failed
    Call FinishReview(id, "Accepted", "Attempt stale")
    Exit Function
Failed:
    StaleReviewRejected = InStr(Err.Description, "stale") > 0
End Function
Private Function CorruptRestoreRejected(ByVal id As Long, ByVal revision As Long) As Boolean
    On Error GoTo Failed
    Call RestoreDraft(id, revision): Exit Function
Failed:
    CorruptRestoreRejected = InStr(Err.Description, "incomplete") > 0
End Function
Public Function TestAdministration() As String
    Dim kind As Variant, id As Long, first As Long, copy As Long, second As Long, values As Variant, bad As Variant
    Dim keys As Variant, limits As Variant, i As Long, before As Long, task As Long, r As DAO.Recordset, recovery As Long
    On Error GoTo Failed
    Call ConfigureWorkspace("SYNTHETIC ADMIN", "Synthetic Reviewer")
    keys = AdministrationKeys(): limits = Array(27, 5, 12, 5, 18)
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        values = Array("O'NAME, SYNTHETIC", "TEST", "0000", "00123", "TEST STATION")
        id = CommitDraft(0, CStr(kind), "Synthetic", "Local display", Null, Null, "Notes", "Administrative draft", "Narrative", values)
        first = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check Stored(id, "OfficialName") = values(0) And Stored(id, "MemberName") = "Local display", "Names combined"
        Check Stored(id, "UnitUIC") = "00123", "Leading zero lost"
        Check DCount("*", "RevisionValue", "RevisionID=" & first) = UBound(DraftFieldKeys()) + 1, "Snapshot not complete"
        Call CommitDraft(id, CStr(kind), "Synthetic", "Local display", Null, Null, "Notes", "Omitted fields", "Narrative")
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1 And Stored(id, "UnitUIC") = "00123", "Omitted administration changed values"
        copy = DuplicateDraft(id)
        Check Stored(copy, "UnitUIC") = "00123" And Stored(copy, "OfficialName") = values(0), "Duplicate lost administration"
        values(0) = "o'name, SYNTHETIC"
        task = AssignReview(id, first, DLookup("ActiveProfileID", "Workspace"), Null, "Review")
        Call CommitDraft(id, CStr(kind), "Synthetic", "Local display", Null, Null, "Notes", "Case", "Narrative", values)
        second = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        Check RevisionFieldStatus(first, second, "OfficialName") = "Changed", "Case edit invisible"
        Check StaleReviewRejected(task), "Administrative edit did not stale review"
        Call FinishReview(task, "Cancelled", "Newer administration")
        Call RestoreDraft(id, first)
        Check Stored(id, "OfficialName") = "O'NAME, SYNTHETIC", "Restore lost administration"
        values = Array(Null, "", Null, "", Null)
        Call CommitDraft(id, CStr(kind), "Synthetic", "Local display", Null, Null, "Notes", "Clear", "Narrative", values)
        Check IsNull(Stored(id, "OfficialName")) And Not IsNull(Stored(id, "UnitUIC")) And Len(Stored(id, "UnitUIC")) = 0, "Null and empty collapsed"
    Next
    before = DCount("*", "ReportRevision")
    For i = 0 To 4
        bad = Array(Null, Null, Null, Null, Null): bad(i) = String$(limits(i) + 1, "X")
        Check Rejected(id, "FITREP", bad), "Overlength accepted for " & keys(i)
    Next
    Check Rejected(id, "FITREP", Array("Name", "TEST", "0000", 123, "Station")), "Numeric UIC accepted"
    Check Rejected(id, "FITREP", Array("Name" & vbCrLf, Null, Null, Null, Null)), "Multiline name accepted"
    Check Rejected(id, "FITREP", Array("Only one")), "Incomplete array accepted"
    Check DCount("*", "ReportRevision") = before, "Rejected administration wrote revision"
    values = Array(String$(27, "N"), String$(5, "R"), String$(12, "D"), "00123", String$(18, "S"))
    Call CommitDraft(id, "FITREP", "Synthetic", "Local display", Null, Null, "Notes", "Boundary", "Narrative", values)
    Check Stored(id, "OfficialName") = values(0) And Stored(id, "ShipStation") = values(4), "Exact boundary truncated"
    before = DCount("*", "ReportRevision")
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<>'UnitUIC'"
    values(3) = "00999": Check Rejected(id, "FITREP", values), "Injected snapshot failure accepted"
    CurrentDb.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = ""
    Check Stored(id, "UnitUIC") = "00123" And DCount("*", "ReportRevision") = before, "Failed save partly committed"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Check Forms!frmEditor!txtUnitUIC = "00123", "Editor did not load administration"
    Forms!frmEditor!txtOfficialName = "UI OFFICIAL NAME": Forms!frmEditor!txtUnitUIC = "00234"
    Check SaveEditor(), "Editor save failed"
    Check Stored(id, "OfficialName") = "UI OFFICIAL NAME" And Stored(id, "UnitUIC") = "00234", "Editor save lost administration"
    Check OpenDraftChecklist(id), "Checklist failed"
    Check Not MissingInput(id, "OfficialName") And Not MissingInput(id, "UnitUIC"), "Complete administration still missing"
    Call CloseDraftChecklist()
    Forms!frmEditor!txtOfficialName = String$(35, "P"): Forms!frmEditor!txtUnitUIC = "00789"
    Check CheckpointEditor(), "Administrative checkpoint failed"
    recovery = DMax("RecoveryDraftID", "RecoveryDraft")
    Check DLookup("UnitUIC", "RecoveryDraft", "RecoveryDraftID=" & recovery) = "00789", "Checkpoint lost UIC"
    Call FixtureCloseEditor()
    Check RestoreRecoveryDraft(recovery), "Administrative recovery failed"
    Check Forms!frmEditor!txtOfficialName = String$(35, "P") And Forms!frmEditor!txtUnitUIC = "00789", "Recovery truncated pending administration"
    Check Stored(id, "UnitUIC") = "00234", "Recovery auto-saved administration"
    Call FixtureCloseEditor()
    'An actual old snapshot lacks both the schema marker and administrative keys.
    first = DMin("RevisionID", "ReportRevision")
    copy = DLookup("ReportID", "ReportRevision", "RevisionID=" & first)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & first & " AND FieldKey In ('OfficialName','RankRate','Designator','UnitUIC','ShipStation')", dbFailOnError
    CurrentDb.Execute "UPDATE ReportRevision SET SnapshotSchema=Null WHERE RevisionID=" & first, dbFailOnError
    Check OpenRevisionComparison(copy, first), "Legacy comparison failed"
    Forms!frmRevisionCompare!chkChangedOnly = False: Call RefreshComparisonFields()
    Forms!frmRevisionCompare!lstCompareFields = "UnitUIC": Call LoadComparisonValues()
    Check Forms!frmRevisionCompare!txtOldValue = "[Not recorded in this revision]", "Legacy marker missing"
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
    Call RestoreDraft(copy, first)
    Check IsNull(Stored(copy, "UnitUIC")), "Legacy restore invented administration"
    second = DLookup("CurrentRevisionID", "Report", "ReportID=" & copy)
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & second & " AND FieldKey='UnitUIC'", dbFailOnError
    Check CorruptRestoreRejected(copy, second), "Incomplete new snapshot restored"
    'Repair the test corruption so migration fixtures retain coherent current history.
    Set r = CurrentDb.OpenRecordset("RevisionValue", dbOpenDynaset)
    r.AddNew: r!RevisionID = second: r!FieldKey = "UnitUIC": r!IsNullValue = True: r.Update: r.Close
    TestAdministration = "PASS: " & checks & " administration assertions": Exit Function
Failed:
    TestAdministration = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function VerifyAdministrativeRecovery() As Boolean
    Dim recovery As Long
    recovery = DMax("RecoveryDraftID", "RecoveryDraft")
    Call RestoreRecoveryDraft(recovery)
    VerifyAdministrativeRecovery = (Forms!frmEditor!txtOfficialName = String$(35, "P") And Forms!frmEditor!txtUnitUIC = "00789")
    Call FixtureCloseEditor()
End Function

Private Sub FixtureCloseEditor()
    ' Preserve the old abrupt-close semantics for persistence/recovery fixtures.
    ' Real save/discard/cancel exits are exercised by UXWorkflowHarness.
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Sub
