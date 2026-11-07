Attribute VB_Name = "RevisionRetentionHarness"
Option Compare Binary
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal text As String)
    If Not ok Then Err.Raise vbObjectError + 9890, , text
    checks = checks + 1
End Sub
Private Sub Edit(ByVal id As Long, ByVal key As String, ByVal value As Variant)
    Dim fields As Object
    Set fields = CreateObject("Scripting.Dictionary"): fields.Add key, value
    Call CommitSavedFields(id, CurrentRevision(id), fields, "Retention test")
End Sub
Private Function SaveRejected(ByVal id As Long, ByVal expected As Long, ByVal value As String, Optional ByVal reason As String = "Retention test") As Boolean
    Dim fields As Object
    Set fields = CreateObject("Scripting.Dictionary"): fields.Add "Title", value
    On Error GoTo Rejected
    Call CommitSavedFields(id, expected, fields, reason)
    Exit Function
Rejected:
    SaveRejected = True
End Function
'Deleting a working state is safe only if every revision pointer is recognised.
'Fails on a relationship to ReportRevision.RevisionID, or a whole-number column
'named like a revision, that RevisionReferenceField does not cover.
Public Function TestRevisionReferenceCoverage() As String
    Dim db As DAO.Database, rel As DAO.Relation, f As DAO.Field, t As DAO.TableDef, missed As String, n As Long
    Set db = CurrentDb
    For Each rel In db.Relations
        If rel.Table = "ReportRevision" Then
            For Each f In rel.Fields
                If f.Name = "RevisionID" Then
                    n = n + 1
                    If Not RevisionOwnedTable(rel.ForeignTable) And Not RevisionReferenceField(rel.ForeignTable, f.ForeignName) Then missed = missed & " " & rel.ForeignTable & "." & f.ForeignName
                End If
            Next
        End If
    Next
    For Each t In db.TableDefs
        If Left$(t.Name, 4) <> "MSys" And Left$(t.Name, 1) <> "~" And Not RevisionOwnedTable(t.Name) Then
            For Each f In t.Fields
                If f.Type = dbLong And InStr(1, f.Name, "Revision", vbTextCompare) > 0 Then
                    Select Case t.Name & "." & f.Name
                        'Not pointers: the revision's own key and number, a count, and a report key.
                        Case "ReportRevision.RevisionID", "ReportRevision.RevisionNumber", "DraftDisposal.RevisionsCleared", "GroupCompletionReport.RevisionOwnerReportID"
                        Case Else
                            n = n + 1
                            If Not RevisionReferenceField(t.Name, f.Name) Then missed = missed & " " & t.Name & "." & f.Name
                    End Select
                End If
            Next
        End If
    Next
    If Len(missed) > 0 Then TestRevisionReferenceCoverage = "FAIL: revision pointers not protected from deletion:" & missed: Exit Function
    TestRevisionReferenceCoverage = "PASS: " & n & " revision relationships and columns are protected from working-state deletion"
End Function
Public Function TestRevisionRetention() As String
    Dim id As Long, first As Long, working As Long, kept As Long, count As Long, kind As Variant, stage As String
    Dim template As Long, run As Long, message As Long, fields As Object, key As String, rule As String, pass As Long
    On Error GoTo Failed
    If DCount("*", "Workspace") = 0 Then ConfigureWorkspace "SYNTHETIC RETENTION", "Tester"
    Check UXLeaveEditor("Discard"), "Could not close editor fixture"
    For Each kind In Array("EVAL", "FITREP", "CHIEFEVAL")
        stage = CStr(kind) & " ordinary saves"
        id = CommitDraft(0, CStr(kind), "Retention " & kind, "Synthetic retention member", Null, Null, "")
        first = CurrentRevision(id)
        Edit id, "Title", "Administrative one": working = CurrentRevision(id)
        Edit id, "Title", "Administrative two"
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 2, "Routine saves accumulated snapshots"
        Check DCount("*", "ReportRevision", "RevisionID=" & working) = 0, "Superseded working snapshot survived"
        Check SaveRejected(id, working, "Stale overwrite"), "Replaced working state did not reject a stale save"
        Check DCount("*", "AuditEvent", "SourceKind='Revision' AND ReportID=" & id) = 1, "Routine saves accumulated audit revisions"
        Check DCount("*", "ReportRevision", "ReportID=" & id & " AND " & HistoryRevisionPredicate("ReportRevision")) = 1, "Working state shown in history"
        AssertReportCache id
        stage = CStr(kind) & " narrative"
        Edit id, "OfficialComments", "First narrative": kept = CurrentRevision(id)
        Check RevisionRetentionKind(kept) = "Narrative", "Main narrative save not retained"
        Edit id, "Title", "Administrative three": working = CurrentRevision(id)
        Edit id, "Title", "Administrative four"
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 3, "Narrative plus working state count wrong"
        Check DCount("*", "ReportRevision", "RevisionID=" & kept) = 1, "Routine edit removed narrative history"
        Edit id, "OfficialComments", ""
        Check RevisionRetentionKind(CurrentRevision(id)) = "Narrative", "Intentional narrative clear not retained"
        If kind = "CHIEFEVAL" Then
            Edit id, "Trait33Comment1", "Chief authored comment"
            Check RevisionRetentionKind(CurrentRevision(id)) = "Narrative", "Chief trait comment not retained"
        End If
        Edit id, "DutyDescription", "Duty narrative"
        Check RevisionRetentionKind(CurrentRevision(id)) = "Narrative", "Duty narrative not retained"
        Edit id, "CommandEmployment", "Command narrative"
        Check RevisionRetentionKind(CurrentRevision(id)) = "Narrative", "Employment narrative not retained"
        stage = CStr(kind) & " checkpoint"
        Edit id, "Title", "Checkpoint state": working = CurrentRevision(id)
        RetainReportCheckpoint id, "Review boundary"
        Edit id, "Title", "After checkpoint"
        Check RevisionRetentionKind(working) = "Checkpoint", "Checkpoint was replaced"
        Set fields = RevisionFields(working): Check fields("Title") = "Checkpoint state", "Checkpoint contents changed"
        stage = CStr(kind) & " referenced snapshot"
        working = CurrentRevision(id)
        message = AddReportMessage(id, "Synthetic context", 0, working)
        Edit id, "Title", "After message"
        Check RevisionRetentionKind(working) = "Protected", "Referenced working state was not protected"
        Set fields = RevisionFields(working): Check fields("Title") = "After checkpoint", "Conversation context changed"
        Check DCount("*", "ReportRevision", "RevisionID=" & working & " AND " & HistoryRevisionPredicate("ReportRevision")) = 0, "Technical context appeared in checkpoint history"
        AssertReportCache id
    Next
    stage = "legacy preservation and rollback"
    id = CommitDraft(0, "EVAL", "Legacy fixture", "Synthetic legacy", Null, Null, "")
    Edit id, "Title", "Legacy saved state": kept = CurrentRevision(id)
    CurrentDb.Execute "DELETE FROM RevisionRetention WHERE RevisionID=" & kept, dbFailOnError
    Edit id, "Title", "Working one": Edit id, "Title", "Working two"
    Set fields = RevisionFields(kept): Check fields("Title") = "Legacy saved state", "Pre-policy history was pruned or rewritten"
    Check DCount("*", "ReportRevision", "RevisionID=" & kept & " AND " & HistoryRevisionPredicate("ReportRevision")) = 0, "Legacy administrative save cluttered checkpoint history"
    working = CurrentRevision(id): count = DCount("*", "ReportRevision")
    rule = CurrentDb.TableDefs("ReportRevision").Fields("Reason").ValidationRule
    CurrentDb.TableDefs("ReportRevision").Fields("Reason").ValidationRule = "<> 'Retention late failure'"
    Check SaveRejected(id, working, "Rejected write", "Retention late failure"), "Injected late revision failure was accepted"
    CurrentDb.TableDefs("ReportRevision").Fields("Reason").ValidationRule = rule
    Check CurrentRevision(id) = working And DCount("*", "ReportRevision") = count, "Failed save pruned or advanced working state"
    Set fields = SavedFields(id): Check fields("Title") = "Working two", "Failed save changed current report"
    stage = "explicit audit operation"
    message = BeginOperation("Retention audit operation", "Synthetic")
    Check AuditRecord("Revision", working, message) > 0, "Explicit revision operation lost its audit envelope"
    Check RevisionRetentionKind(working) = "Protected", "Explicit operation snapshot not protected"
    FinishOperation message, "Completed", "Synthetic"
    stage = "workflow boundaries"
    template = CreateRouteTemplate("Retention route", CStr(ActiveContributorID()), "Synthetic route")
    run = StartRoute(id, template)
    Check RevisionRetentionKind(CurrentRevision(id)) = "Checkpoint", "Route start not checkpointed"
    Edit id, "Title", "Before forward": working = CurrentRevision(id)
    ForwardRoute run, "Synthetic forward"
    Check RevisionRetentionKind(working) = "Checkpoint", "Forward not checkpointed"
    Edit id, "Title", "Before return": working = CurrentRevision(id)
    SendBackRoute run, 0, "Synthetic return"
    Check RevisionRetentionKind(working) = "Checkpoint", "Return not checkpointed"
    ForwardRoute run, "Resubmit": ForwardRoute run, "Forward again"
    ApproveSavedRoute run, "Synthetic approval"
    Check RevisionRetentionKind(CurrentRevision(id)) = "Checkpoint", "Approval not checkpointed"
    Call MarkReportSigned(id, CurrentRevision(id), Date)
    Check RevisionRetentionKind(CurrentRevision(id)) = "Checkpoint", "Signature not checkpointed"
    stage = "editor recovery references"
    id = CommitDraft(0, "EVAL", "Recovery retention", "Synthetic recovery", Null, Null, "")
    Edit id, "Title", "Before UI"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Check Forms!frmEditor!uxFieldSurface.Form!uxCounselorPicker.ControlType = acCommandButton, "Counselor picker is not a button"
    Check Forms!frmEditor!uxFieldSurface.Form!uxCounselorPicker.Visible, "Counselor button not visible"
    For pass = 1 To 3
        working = CurrentRevision(id)
        Forms!frmEditor!txtTitle = "Recovery save " & pass: UXSurfaceRefresh
        Check CheckpointEditor(), "Could not create pending recovery checkpoint"
        Check UXSaveEditor(), "Editor save with recovery failed"
        Check DCount("*", "ReportRevision", "RevisionID=" & working) = 0, "Cleared recovery reference retained an old working snapshot"
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 2, "Recovery checkpoints caused routine history accumulation"
    Next
    Check UXLeaveEditor("Discard"), "Could not close recovery fixture"
    stage = "stale working-state comparison"
    id = CommitDraft(0, "EVAL", "Comparison retention", "Synthetic comparison", Null, Null, "")
    first = CurrentRevision(id): Edit id, "OfficialName", "FIRST, STATE": working = CurrentRevision(id)
    Check OpenRevisionComparison(id, first, working), "Working-state comparison failed"
    Forms!frmRevisionCompare!lstCompareFields = "OfficialName"
    Edit id, "OfficialName", "SECOND, STATE"
    Check Not LoadComparisonValues(), "Stale comparison silently showed an empty replaced state"
    Check Forms!frmRevisionCompare!txtOldValue = "" And Forms!frmRevisionCompare!txtNewValue = "", "Stale comparison left misleading values"
    Check InStr(Forms!frmRevisionCompare!lblSummary.Caption, "reopen comparison") > 0, "Stale comparison did not explain recovery"
    Check Not RefreshComparisonFields(), "Stale full comparison was accepted"
    Check OpenRevisionComparison(id, first, CurrentRevision(id)), "Fresh comparison retry failed"
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
    TestRevisionRetention = "PASS: " & checks & " working-state, narrative, checkpoint, context and routing retention assertions": Exit Function
Failed:
    TestRevisionRetention = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
    On Error Resume Next
    If stage = "legacy preservation and rollback" Then CurrentDb.TableDefs("ReportRevision").Fields("Reason").ValidationRule = rule
End Function
