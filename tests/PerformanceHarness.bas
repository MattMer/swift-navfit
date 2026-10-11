Attribute VB_Name = "PerformanceHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 7990, , message
    checks = checks + 1
End Sub
Public Function TestPerformance() As String
    Dim id As Long, groupID As Long, first As Long, second As Long, third As Long, changes As Object, checkpoint As Date, f As Form, caption As String
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC PERFORMANCE", "Tester"
    id = CommitDraft(0, "EVAL", "Synthetic", "Synthetic", Null, Null, "", "Test", Null, , Array("4.0", "4.0", "4.0", "4.0", "4.0", "4.0", "4.0"), Array("Observed", "Promotable"))
    groupID = CreateSummaryGroup("Performance group", "EVAL", "E5")
    AddGroupMember groupID, id, "E5"
    first = ReuseGroupComparison(groupID, "Comparison")
    Check ReuseGroupComparison(groupID, "Comparison") = first, "Unchanged baseline duplicated"
    Check DCount("*", "GroupSnapshot") = 1 And DCount("*", "ComparisonMember") = 1, "Navigation wrote history"
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "Trait33", "3.0"
    CommitSavedFields id, CurrentRevision(id), changes, "Test changed revision"
    second = ReuseGroupComparison(groupID, "Comparison")
    Check second <> first, "Changed revision reused stale baseline"
    Check DLookup("RevisionID", "ComparisonMember", "SetID=" & first) <> CurrentRevision(id), "Historical baseline mutated"
    ProposeMeetingField second, id, "Trait34", "3.0"
    third = ReuseGroupComparison(groupID, "Comparison")
    Check third <> second, "Proposal baseline reused for navigation"
    Check ReuseGroupComparison(groupID, "Comparison") = third, "Fresh proposal-free baseline not reused"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Set f = Forms!frmEditor
    f!txtTitle = "Pending performance"
    Check CheckpointEditor(), "Initial checkpoint failed"
    CurrentDb.Execute "UPDATE RecoveryDraft SET CheckpointAt=#1/1/2000#", dbFailOnError
    Check CheckpointEditor(), "Unchanged checkpoint failed"
    Check DLookup("CheckpointAt", "RecoveryDraft") = #1/1/2000#, "Unchanged checkpoint rewritten"
    f!txtTitle = "Changed performance"
    Check CheckpointEditor(), "Changed checkpoint failed"
    Check DLookup("DraftTitle", "RecoveryDraft") = "Changed performance" And DLookup("CheckpointAt", "RecoveryDraft") <> #1/1/2000#, "Changed checkpoint skipped"
    CurrentDb.Execute "DELETE FROM RecoveryDraft", dbFailOnError
    Check CheckpointEditor() And DCount("*", "RecoveryDraft") = 1, "Deleted checkpoint not restored"
    Check UXLivePaint(), "Validation failed"
    caption = f!uxDirty.Caption
    Check UXLivePaint() And f!uxDirty.Caption = caption, "Repeated validation changed findings"
    f!txtMemberSSN = "INVALID"
    Check UXLivePaint(), "Changed validation failed"
    Check InStr(f!uxDirty.ControlTipText, "invalid entries") > 0, "Validation status missing"
    TestPerformance = "PASS: " & checks & " baseline reuse, immutable history, proposal isolation, checkpoint and validation assertions"
    Exit Function
Failed:
    TestPerformance = "FAIL: " & Err.Number & " " & Err.Description
End Function
