Attribute VB_Name = "UXMemberGridHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6702, , message
    n = n + 1
End Sub
Public Function TestUXMemberGrid() As String
    Dim id As Long, g As Long, emptyGroup As Long, rev As Long, rows As Access.Form
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC MEMBER GRID", "Tester"
    id = CommitDraft(0, "EVAL", "Member grid", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test", "")
    g = CreateSummaryGroup("Member grid", "EVAL", "E5"): AddGroupMember g, id, "E5"
    DoCmd.OpenForm "frmUXGroups": Forms!frmUXGroups!lstGroups = g
    Check UXOpenGroupMembers(), "Group did not open member grid"
    Check CurrentProject.AllForms("frmUXComparison").IsLoaded, "Default member grid absent"
    Set rows = Forms!frmUXComparison!uxTraitGrid.Form
    Check rows!uxOpenReport.OnClick = "=UXGridOpenReport()", "Row Open action missing"
    Check UXGridOpenReport(), "Row report entry failed"
    Check Forms!frmEditor!txtReportID = id, "Row opened wrong report"
    rev = CurrentRevision(id)
    Forms!frmEditor!txtOfficialComments = "Edited from member grid": UXSurfaceRefresh
    Check UXFeedbackCloseEditor(), "Report return to members failed"
    Check Screen.ActiveForm.Name = "frmUXComparison", "Return did not select member grid"
    Check rows!ReportID = id, "Return lost selected member"
    Check CurrentRevision(id) <> rev, "Report return did not save pending values"
    Check rows!AssignmentRecommendationA.Locked And rows!AssignmentRecommendationB.Locked, "Stale career fields not visibly locked"
    UXCloseComparison
    emptyGroup = CreateSummaryGroup("Empty group", "EVAL", "E5")
    Forms!frmUXGroups!lstGroups = emptyGroup
    Check UXOpenGroupMembers(), "Empty group member grid failed"
    Check Not CurrentProject.AllForms("frmUXComparison").IsLoaded And Forms!frmMajorWorkbench!lstReports.ListCount = 0, "Empty group displayed unrelated members"
    TestUXMemberGrid = "PASS: " & n & " member-grid entry and return assertions": Exit Function
Failed:
    TestUXMemberGrid = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
