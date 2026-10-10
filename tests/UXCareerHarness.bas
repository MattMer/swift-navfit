Attribute VB_Name = "UXCareerHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6701, , message
    n = n + 1
End Sub
Public Function TestUXCareers() As String
    Dim id As Long, g As Long, rev As Long, rows As Access.Form, f As Access.Form, values As Object
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC CAREER", "Tester"
    id = CommitDraft(0, "EVAL", "Career", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test", "")
    g = CreateSummaryGroup("Career group", "EVAL", "E5"): AddGroupMember g, id, "E5"
    DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = g
    Check UXOpenComparison(), "Comparison entry failed"
    Set f = Forms!frmUXComparison: Set rows = f!uxTraitGrid.Form: rev = CurrentRevision(id)
    rows!AssignmentRecommendationA = "INSTRUCTOR"
    Check UXGridEdit("AssignmentRecommendationA"), "First career field not staged"
    rows!AssignmentRecommendationB = "RECRUITER"
    Check UXGridEdit("AssignmentRecommendationB"), "Second career save failed: " & f!lblStatus.Caption
    Check CurrentRevision(id) <> rev, "Career cells did not save automatically"
    Check DCount("*", "MeetingProposal", "SetID=" & f!txtMeeting & " AND ProposalState='Proposed'") = 0, "Automatic save left pending proposals"
    rows!AssignmentRecommendationA = String$(21, "X")
    Check Not UXGridEdit("AssignmentRecommendationA"), "Oversize recommendation accepted"
    Check UXGridLoad(), "Career cache reload failed"
    Set rows = f!uxTraitGrid.Form
    Check rows!AssignmentRecommendationA = "INSTRUCTOR" And rows!AssignmentRecommendationB = "RECRUITER", "Pending career recommendations lost on reload"
    Check DLookup("AssignmentRecommendationA", "Report", "ReportID=" & id) = "INSTRUCTOR" And DLookup("AssignmentRecommendationB", "Report", "ReportID=" & id) = "RECRUITER", "Applied career values wrong"
    UXCloseComparison
    Check UXTableITA(id) = "Unavailable", "Missing traits implied ITA"
    Check UXTableGroups(id) = "Career group", "Report group label missing"
    DoCmd.OpenForm "frmLibrary": RefreshLibrary
    Check Forms!frmLibrary!lstReports.ColumnCount = 9, "Report table columns missing"
    Check UXWorkbenchMode("Members"), "Group member view failed"
    Check Forms!frmMajorWorkbench!lstReports.ColumnCount = 7, "Group report columns missing"
    DoCmd.OpenForm "frmUXGroups": Forms!frmUXGroups!lstGroups = g
    Check UXListGroupDefaults(), "Group-list defaults entry failed"
    Check Forms!frmUXBulk!uxBlocks.Visible, "Defaults section editor absent"
    TestUXCareers = "PASS: " & n & " career and table assertions": Exit Function
Failed:
    TestUXCareers = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
