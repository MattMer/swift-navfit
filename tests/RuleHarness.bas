Attribute VB_Name = "RuleHarness"
Option Compare Binary
Option Explicit
Public Function TestRules() As String
    Dim size As Long, result As Object, sixty As Variant, fifty As Variant, forty As Variant, count As Long, pack As Long, id As Long, run As Long, changes As Object
    On Error GoTo Failed
    ' Independently transcribed from the supplied instruction's rendered table 1-2.
    sixty = Array(0, 1, 1, 2, 2, 2, 3, 3, 4, 4, 4, 5, 5, 6, 6, 6, 7, 7, 8, 8, 8, 9, 9, 10, 10, 10, 11, 11, 12, 12)
    fifty = Array(0, 1, 1, 1, 2, 1, 2, 2, 3, 3, 3, 3, 4, 4, 5, 4, 5, 5, 6, 6, 6, 6, 7, 7, 8, 7, 8, 8, 9, 9)
    forty = Array(0, 1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 4, 4, 4, 4, 4, 5, 5, 5, 5, 5, 6, 6, 6)
    For size = 1 To 30
        Set result = PromotionCapacity("E5", size, False)
        If result("MustPromoteMaximum") <> sixty(size - 1) Then Err.Raise 5, , "60-percent table mismatch at size " & size
        Set result = PromotionCapacity("E7", size, False)
        If result("MustPromoteMaximum") <> fifty(size - 1) Then Err.Raise 5, , "50-percent table mismatch at size " & size
        Set result = PromotionCapacity("O5", size, False)
        If result("MustPromoteMaximum") <> forty(size - 1) Then Err.Raise 5, , "40-percent table mismatch at size " & size
        count = count + 3
    Next
    Set result = PromotionCapacity("E6", 42, False)
    If result("EarlyPromoteMaximum") <> 9 Or result("MustPromoteMaximum") <> 17 Then Err.Raise 5, , "Instruction's 42-member example mismatch"
    count = count + 1
    Set result = PromotionCapacity("E7", 6, False, 1)
    If result("MustPromoteMaximum") <> 2 Then Err.Raise 5, , "Unused EP transfer failed"
    count = count + 1
    Set result = PromotionCapacity("O1", 2, False)
    If result("CombinedMaximum") <> 0 Then Err.Raise 5, , "Non-LDO prohibition must prevail"
    count = count + 1
    Set result = PromotionCapacity("O1", 2, True)
    If result("CombinedMaximum") <> 2 Then Err.Raise 5, , "LDO exception failed"
    count = count + 1
    Call ConfigureWorkspace("SYNTHETIC COMMAND", "Synthetic rule tester")
    pack = InstallLegacyNavyDraftChecks()
    If InstallLegacyNavyDraftChecks() <> pack Or DCount("*", "RulePack") <> 1 Then Err.Raise 5, , "Legacy rule pack installation is not idempotent"
    count = count + 1
    id = CommitDraft(0, "EVAL", "Synthetic policy checks", "Synthetic Example", #1/1/2026#, #9/30/2026#, "", "Test")
    run = RunRulePack(id, pack)
    If DCount("*", "ValidationFinding", "ValidationRunID=" & run & " AND FindingState='Fail'") <> 3 Then Err.Raise 5, , "Required-field findings missing"
    count = count + 1
    Set changes = CreateObject("Scripting.Dictionary")
    changes.Add "PromotionRecommendation", "Early Promote": changes.Add "Trait33", "2.0"
    Call CommitSavedFields(id, CurrentRevision(id), changes, "Test constrained recommendation")
    run = RunRulePack(id, pack)
    If DCount("*", "ValidationFinding", "ValidationRunID=" & run & " AND FindingState='Fail'") <> 4 Then Err.Raise 5, , "Trait recommendation constraint missing"
    count = count + 1
    TestRules = "PASS: " & count & " source-table and rule assertions": Exit Function
Failed:
    TestRules = "FAIL after " & count & ": " & Err.Number & " / " & Err.Description
End Function
