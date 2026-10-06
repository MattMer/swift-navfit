Attribute VB_Name = "UXGlobalDefaultsHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6900, , message
    n = n + 1
End Sub
Public Function TestUXGlobalDefaults() As String
    Dim id As Long, g As Long, rev As Long, result As Object
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC GLOBAL DEFAULTS", "Tester"
    id = CommitDraft(0, "EVAL", "Existing", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test", "")
    g = CreateSummaryGroup("Defaults group", "EVAL", "E5"): AddGroupMember g, id, "E5": rev = CurrentRevision(id)
    Check UXOpenGlobalDefaults(), "Global defaults required a selected group"
    Check Forms!frmUXBulk!lblHeading.Caption = "Global defaults", "Global scope not labeled"
    Forms!frmUXBulk!uxBlocks.Form!valueUnitUIC = "54321": UXBlockEdit "UnitUIC"
    Forms!frmUXBulk!txtReason = "Global starting unit"
    Check UXPreviewBulk(), "Global default preview failed"
    Check InStr(Forms!frmUXBulk!lstPreview.RowSource, "Global default") > 0, "Preview mislabeled scope"
    Check UXApplyBulk(), "Global defaults apply failed"
    Check UXGroupDefaultValue(0, "UnitUIC") = "54321", "Global default not stored"
    Check CurrentRevision(id) = rev, "Global defaults rewrote saved report"
    UXCloseBlocks
    Set result = ResolveFieldDefault(id, "UnitUIC", g)
    Check result("Value") = "54321", "Group did not inherit global suggestion"
    SetFieldDefault "Group", g, "UnitUIC", "12345"
    Set result = ResolveFieldDefault(id, "UnitUIC", g)
    Check result("Value") = "12345", "Group override lost to global"
    SetFieldDefault "Report", id, "UnitUIC", "67890"
    Set result = ResolveFieldDefault(id, "UnitUIC", g)
    Check result("Value") = "67890", "Report override lost to parent"
    DoCmd.OpenForm "frmEditor"
    Check Forms!frmEditor!txtUnitUIC = "54321", "New report did not receive global starting value"
    Forms!frmEditor!txtUnitUIC = "11111"
    Check Forms!frmEditor!txtUnitUIC = "11111", "Inherited value locked editing"
    UXLeaveEditor "Discard"
    Check UXOpenGlobalDefaults(), "Global defaults reopening failed"
    Forms!frmUXBulk!uxBlocks.Form!valueUnitUIC = "22222": UXBlockEdit "UnitUIC"
    Forms!frmUXBulk!txtReason = "Stale test": UXPreviewBulk
    SetFieldDefault "Command", 1, "UnitUIC", "33333"
    Check Not UXApplyBulk(), "Concurrent global edit overwritten"
    Check UXGroupDefaultValue(0, "UnitUIC") = "33333", "Rejected global edit changed value"
    UXCloseBlocks
    DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = g
    Check UXOpenGroupDefaults(), "Group defaults reset entry failed"
    Check Forms!frmUXBulk!uxBlocks.Form!valueUnitUIC = "12345", "Group override not displayed"
    Check UXResetDefault("UnitUIC"), "Reset did not stage inheritance"
    Forms!frmUXBulk!txtReason = "Reset group override"
    Check UXPreviewBulk() And UXApplyBulk(), "Reset group override did not apply"
    Check DCount("*", "FieldDefault", "ScopeKind='Group' AND ScopeID=" & g & " AND FieldKey='UnitUIC'") = 0, "Reset stored null instead of removing override"
    Check UXEffectiveDefault(g, "UnitUIC") = "33333", "Reset group did not inherit global"
    Check CurrentRevision(id) = rev, "Reset defaults changed saved report"
    UXCloseBlocks
    Check UXOpenGlobalDefaults(), "Reset global entry failed"
    Check UXResetDefault("UnitUIC"), "Global reset did not stage"
    Forms!frmUXBulk!txtReason = "Reset global default"
    Check UXPreviewBulk() And UXApplyBulk(), "Global reset failed"
    Check DCount("*", "FieldDefault", "ScopeKind='Command' AND ScopeID=1 AND FieldKey='UnitUIC'") = 0, "Global reset did not remove default"
    TestUXGlobalDefaults = "PASS: " & n & " global default assertions": Exit Function
Failed:
    UXCloseBlocks
    DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = g
    Check UXOpenGroupDefaults(), "Group defaults reset entry failed"
    Check Forms!frmUXBulk!uxBlocks.Form!valueUnitUIC = "12345", "Group override not displayed"
    Check UXResetDefault("UnitUIC"), "Reset did not stage inheritance"
    Forms!frmUXBulk!txtReason = "Reset group override"
    Check UXPreviewBulk() And UXApplyBulk(), "Reset group override did not apply"
    Check DCount("*", "FieldDefault", "ScopeKind='Group' AND ScopeID=" & g & " AND FieldKey='UnitUIC'") = 0, "Reset stored null instead of removing override"
    Check UXEffectiveDefault(g, "UnitUIC") = "33333", "Reset group did not inherit global"
    Check CurrentRevision(id) = rev, "Reset defaults changed saved report"
    UXCloseBlocks
    Check UXOpenGlobalDefaults(), "Reset global entry failed"
    Check UXResetDefault("UnitUIC"), "Global reset did not stage"
    Forms!frmUXBulk!txtReason = "Reset global default"
    Check UXPreviewBulk() And UXApplyBulk(), "Global reset failed"
    Check DCount("*", "FieldDefault", "ScopeKind='Command' AND ScopeID=1 AND FieldKey='UnitUIC'") = 0, "Global reset did not remove default"
    TestUXGlobalDefaults = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
