Attribute VB_Name = "NavigationContextHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8950, , message
 n = n + 1
End Sub
Public Function TestNavigationContext() As String
 Dim a As Long, b As Long, id As Long, before As Long, f As Access.Form, rows As Access.Form
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC NAVIGATION CONTEXT", "Tester"
 a = CreateSummaryGroup("Group Alpha", "EVAL", "E5")
 b = CreateSummaryGroup("Group Bravo", "CHIEFEVAL", "E7")
 id = CreateGroupDraft(a, GroupStamp(a), UXDefaultsStamp(a), 0, 0, "Synthetic", "SYNTHETIC, MEMBER", "IT2", "", Null, Null)
 Check UXOpenGroupDefaults(a), "Direct defaults needs hidden workbench"
 Check Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded, "Defaults created a hidden host"
 Check Forms!frmUXBulk!uxGroupContext = a, "Direct context wrong"
 before = DCount("*", "ReportRevision")
 Check UXGroupNavigate("frmUXBulk", 1), "Repeated defaults tab failed"
 Check DCount("*", "ReportRevision") = before, "Repeated page created revisions"
 Check Not UXGroupNavigate("frmUXBulk", 9), "Invalid page accepted"
 Check Forms!frmUXBulk!uxGroupContext = a, "Invalid page changed context"
 UXCloseBlocks
 DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = b
 Check UXOpenGroupDefaults(a), "Explicit override failed"
 Check Forms!frmUXBulk!uxGroupContext = a And Forms!frmUXBulk!uxKind = "EVAL", "Hidden selection contaminated defaults"
 UXCloseBlocks
 Check UXOpenGlobalDefaults(), "Global defaults failed"
 Check Forms!frmUXBulk!uxGroupContext = 0, "Global scope retained group"
 Check UXOpenGroupDefaults(b), "Global to explicit group failed"
 Check Forms!frmUXBulk!uxGroupContext = b And Forms!frmUXBulk!uxKind = "CHIEFEVAL", "Group switch retained global/type context"
 UXCloseBlocks
 DoCmd.Close acForm, "frmMajorWorkbench", acSaveNo
 DoCmd.OpenForm "frmUXComparison", , , , , , "G:" & a
 Set f = Forms!frmUXComparison: Set rows = f!uxTraitGrid.Form
 rows!AssignmentRecommendationA = String$(21, "X")
 Check UXGridQueue("AssignmentRecommendationA"), "Invalid value not queued"
 Check Not UXOpenGroupDefaults(b), "Failed cell allowed defaults navigation"
 Check UXGridPendingCount() = 1 And f!uxGroupContext = a, "Blocked transition lost pending/context"
 Check Not CurrentProject.AllForms("frmUXBulk").IsLoaded, "Blocked transition opened defaults"
 rows!AssignmentRecommendationA = "VALID": Check UXGridQueue("AssignmentRecommendationA"), "Retry correction failed"
 Check UXGroupNavigate("frmUXComparison", 1), "Member to defaults transition failed"
 Check Forms!frmUXBulk!uxGroupContext = a, "Transition chose another group"
 Check DLookup("AssignmentRecommendationA", "Report", "ReportID=" & id) = "VALID", "Transition did not save corrected cell"
 Check Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded, "Transition created hidden defaults host"
 Check UXGroupNavigate("frmUXBulk", 0), "Return to Members failed"
 Check Forms!frmUXComparison!uxGroupContext = a, "Return lost group"
 Check UXComparisonPage("Preview"), "Proposal page failed"
 Check Not Forms!frmUXComparison!uxTraitGrid.Visible And Forms!frmUXComparison!lstProposals.Visible, "Proposal visibility wrong"
 Check UXComparisonPage("Marks"), "Marks page failed"
 Check Forms!frmUXComparison!uxTraitGrid.Visible And Not Forms!frmUXComparison!lstProposals.Visible, "Marks visibility wrong"
 Check Not Forms!frmUXComparison!cmdStage.Visible, "Retired marks action reappeared"
 Check UXCloseComparison(False), "Close members failed"
 TestNavigationContext = "PASS: " & n & " explicit group context, pending-cell transition and page visibility assertions": Exit Function
Failed:
 TestNavigationContext = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
