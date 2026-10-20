Attribute VB_Name = "GroupExitHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8970, , message
 n = n + 1
End Sub
Private Sub OpenGroup(ByVal id As Long)
 Forms!frmUXGroups!lstGroups = id
 Check UXOpenGroup(), "Open selected group failed"
End Sub
Public Function TestGroupExit() As String
 Dim a As Long, b As Long, emptyGroup As Long, report As Long, other As Long, rows As Access.Form, key As Variant, before As Long
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC GROUP EXIT", "Tester"
 a = CreateSummaryGroup("First group", "EVAL", "E5"): b = CreateSummaryGroup("Second group", "EVAL", "E6")
 emptyGroup = CreateSummaryGroup("Empty group", "EVAL", "E5")
 report = CommitDraft(0, "EVAL", "First member", "Member A", Null, Null, "")
 other = CommitDraft(0, "EVAL", "Second member", "Member B", Null, Null, "")
 AddGroupMember a, report, "E5": AddGroupMember b, other, "E6"
 DoCmd.OpenForm "frmUXGroups"
 OpenGroup a: Check Forms!frmUXComparison!uxGroupContext = a, "First group context wrong"
 Check Forms!frmUXComparison!uxBackToGroups.OnClick = "=UXBackToGroups()", "Back binding missing"
 OpenGroup b: Check Forms!frmUXComparison!uxGroupContext = b And Forms!frmUXComparison!uxTraitGrid.Form!ReportID = other, "Second open retained first group"
 Check UXBackToGroups(), "Back failed"
 Check Not CurrentProject.AllForms("frmUXComparison").IsLoaded And Screen.ActiveForm.Name = "frmUXGroups", "Back left group tab open"
 For Each key In Array("frmHome", "frmLibrary", "frmMyWork", "frmUXResources", "frmUXSettings")
  OpenGroup a
  Check UXNavigate(CStr(key)), "Menu navigation failed: " & key
  Check Not CurrentProject.AllForms("frmUXComparison").IsLoaded And Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded, "Menu retained group tab"
  Check UXNavigate("frmUXGroups"), "Group list return failed"
 Next
 OpenGroup a: Set rows = Forms!frmUXComparison!uxTraitGrid.Form
 rows!AssignmentRecommendationA = String$(21, "X"): Check UXGridQueue("AssignmentRecommendationA"), "Queue failed"
 before = CurrentRevision(report)
 Check Not UXNavigate("frmHome"), "Invalid pending cell allowed leave"
 Check CurrentProject.AllForms("frmUXComparison").IsLoaded And UXGridPendingCount() = 1 And CurrentRevision(report) = before, "Blocked leave lost pending data"
 Forms!frmUXGroups!lstGroups = b: Check Not UXOpenGroup(), "Invalid pending cell allowed replacement"
 Check Forms!frmUXComparison!uxGroupContext = a, "Blocked replacement changed context"
 rows!AssignmentRecommendationA = "CORRECTED": Check UXGridQueue("AssignmentRecommendationA"), "Correction queue failed"
 Check UXBackToGroups(), "Corrected pending exit failed"
 Check DLookup("AssignmentRecommendationA", "Report", "ReportID=" & report) = "CORRECTED", "Exit did not save valid pending cell"
 OpenGroup a
 Check UXGroupNavigate("frmUXComparison", 1), "Defaults transition failed"
 Check UXBackToGroups(), "Defaults Back failed"
 Check Not CurrentProject.AllForms("frmUXBulk").IsLoaded And Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded, "Defaults host retained"
 OpenGroup a
 Check UXGroupNavigate("frmUXComparison", 2), "Review transition failed"
 Check UXNavigate("frmLibrary"), "Review menu leave failed"
 Check Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded, "Review host retained"
 OpenGroup emptyGroup
 Check CurrentProject.AllForms("frmMajorWorkbench").IsLoaded, "Empty group missing"
 Check UXBackToGroups(), "Empty group Back failed"
 Check Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded, "Empty group retained"
 OpenGroup b: Check Forms!frmUXComparison!uxGroupContext = b, "Final group context wrong"
 Check UXGridOpenReport(), "Member editor open failed"
 Forms!frmEditor!txtOfficialComments = "Pending report edit"
 Check UXNavigate("frmHome"), "Nested report/group navigation failed"
 Check Not CurrentProject.AllForms("frmEditor").IsLoaded And Not CurrentProject.AllForms("frmUXComparison").IsLoaded, "Nested navigation retained tabs"
 Check DLookup("OfficialComments", "Report", "ReportID=" & other) = "Pending report edit", "Nested navigation lost report save"
 TestGroupExit = "PASS: " & n & " group replacement, Back, menu exit and pending edit assertions": Exit Function
Failed:
 TestGroupExit = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
