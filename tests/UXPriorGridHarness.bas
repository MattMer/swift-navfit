Attribute VB_Name = "UXPriorGridHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6702, , message
    n = n + 1
End Sub
Public Function TestUXPriorGrid() As String
    Dim id As Long, g As Long, emptyGroup As Long, rev As Long, rows As Access.Form, m As Long, marks As Object, b As Long, prior As Long, changes As Object
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC MEMBER GRID", "Tester"
    id = CommitDraft(0, "EVAL", "Member grid", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test", "")
    m = CLng(DLookup("MemberID", "MemberReport", "ReportID=" & id))
    LinkReportDirectoryMember id, m, CurrentRevision(id), CLng(DLookup("VersionCounter", "CommandMember", "MemberID=" & m)), "Synthetic confirmed identity"
    Set marks = CreateObject("Scripting.Dictionary")
    For b = 33 To 39: marks.Add "Trait" & b, "4.0": Next
    prior = SavePriorEvaluation(0, 0, m, "EVAL", "IT2", #1/1/2025#, #9/30/2025#, "Promotable", "Observed", marks, "Synthetic check-in evaluation")
    g = CreateSummaryGroup("Member grid", "EVAL", "E5"): AddGroupMember g, id, "E5"
    DoCmd.OpenForm "frmUXGroups": Forms!frmUXGroups!lstGroups = g
    Check UXOpenGroupMembers(), "Group did not open member grid"
    Check CurrentProject.AllForms("frmUXComparison").IsLoaded, "Default member grid absent"
    Set rows = Forms!frmUXComparison!uxTraitGrid.Form
    Check rows!PriorITA = "4.00" And rows!PriorPromotion = "Promotable", "Prior values missing from member grid"
    Check rows!PriorITA.Locked And rows!PriorPromotion.Locked, "Historical values editable"
    Check Not rows!PriorITA.Enabled And Not rows!PriorPromotion.Enabled And Not rows!PriorITA.TabStop And Not rows!PriorPromotion.TabStop, "Prior cells can receive focus"
    Check rows!PriorITA.BackColor = UXColor("Neutral") And rows!PriorPromotion.BackColor = UXColor("Neutral"), "Prior cells lack read-only shading"
    Check rows!Trait33.Enabled And rows!PromotionRecommendation.Enabled, "Current-value cells were disabled"
    Check rows!uxOpenReport.OnClick = "=UXGridOpenReport()", "Row Open action missing"
    Check UXGridOpenReport(), "Row report entry failed"
    Check Forms!frmEditor!txtReportID = id, "Row opened wrong report"
    rev = CurrentRevision(id)
    Forms!frmEditor!txtOfficialComments = "Edited from member grid": UXSurfaceRefresh
    Check UXFeedbackCloseEditor(), "Report return to members failed"
    Check Screen.ActiveForm.Name = "frmUXComparison", "Return did not select member grid"
    Check rows!ReportID = id, "Return lost selected member"
    Check CurrentRevision(id) <> rev, "Report return did not save pending values"
    ' Returning reloads the member row at the saved revision, so its cells stay editable.
    Check rows!RevisionID = CurrentRevision(id) And Not rows!AssignmentRecommendationA.Locked And Not rows!AssignmentRecommendationB.Locked, "Return did not refresh the member row"
    ' A save made outside the open grid leaves the row stale; its career cells lock.
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "Title", "Changed outside the grid"
    CommitSavedFields id, CurrentRevision(id), changes, "Saved outside the member grid"
    Check UXGridCurrent(True), "Row lock refresh failed"
    Check rows!AssignmentRecommendationA.Locked And rows!AssignmentRecommendationB.Locked, "Stale career fields not visibly locked"
    UXCloseComparison
    emptyGroup = CreateSummaryGroup("Empty group", "EVAL", "E5")
    Forms!frmUXGroups!lstGroups = emptyGroup
    Check UXOpenGroupMembers(), "Empty group member grid failed"
    ' An empty group opens the same Members surface with no rows and an explicit status.
    Check Forms!frmUXComparison!uxGroupContext = emptyGroup And Forms!frmUXComparison!uxTraitGrid.Form.Recordset.EOF, "Empty group displayed unrelated members"
    Check InStr(Forms!frmUXComparison!lblStatus.Caption, "No reports in this summary group") = 1, "Empty group status missing"
    TestUXPriorGrid = "PASS: " & n & " member-grid entry and return assertions": Exit Function
Failed:
    TestUXPriorGrid = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
