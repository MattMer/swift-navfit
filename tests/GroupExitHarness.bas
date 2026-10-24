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
 Check Not UXHasControl(Forms!frmUXComparison, "uxBackToGroups") And Forms!frmUXComparison!cmdBack.OnClick = "=UXBackToGroups()", "Duplicate header Back remains or bottom Back binding is missing"
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
 Check CurrentProject.AllForms("frmUXComparison").IsLoaded And Forms!frmUXComparison!uxGroupContext = emptyGroup, "Empty group did not use canonical Members page"
 Check UXBackToGroups(), "Empty group Back failed"
 Check Not CurrentProject.AllForms("frmUXComparison").IsLoaded And Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded, "Empty group retained"
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

'Read-only tools must tolerate different task orders without saving pending data.
Public Function TestGroupTaskOrders() As String
 Dim order As Variant, action As Variant, id As Long, g As Long, before As Long, start As Long, stage As String
 Dim rows As Access.Form, capture As Object, reports As Collection
 On Error GoTo Failed
 start = n
 For Each order In Array("PDF,Export,Findings", "PDF,Findings,Export", "Export,PDF,Findings", "Export,Findings,PDF", "Findings,PDF,Export", "Findings,Export,PDF")
  id = CommitDraft(0, "EVAL", "Task order", "SYNTHETIC ORDER", Null, Null, "")
  g = CreateSummaryGroup("Synthetic task order " & order, "EVAL", "E5"): AddGroupMember g, id, "E5"
  Check UXShowGroupMembers(g), "Task-order Members failed"
  Set rows = Forms!frmUXComparison!uxTraitGrid.Form
  rows!AssignmentRecommendationA = String$(21, "X")
  Check UXGridQueue("AssignmentRecommendationA"), "Task-order queue failed"
  before = CurrentRevision(id)
  For Each action In Split(order, ",")
   stage = order & " / " & action
   Select Case action
    Case "PDF"
     Check PDFOutputOpenGroup("frmUXComparison"), "Pending cell blocked PDF preview"
     Set capture = PDFOutputJob(): Set reports = capture("Reports")
     Check reports(1)("Fields")("AssignmentRecommendationA") = String$(21, "X"), "PDF preview omitted pending cell"
     Check PDFOutputClose(), "PDF preview close failed"
    Case "Export"
     Check UXCurrentGroupTool("frmUXComparison", "Export"), "Pending cell blocked saved NAVFIT export dialog"
     Check Forms!frmNAVFITExchange!cboGroup = g, "NAVFIT dialog targeted stale group"
     DoCmd.Close acForm, "frmNAVFITExchange", acSaveNo
    Case "Findings"
     Check UXOpenValidation(True, g), "Read-only findings failed"
     Check Forms!frmUXValidation.OpenArgs = "G:" & g, "Findings targeted stale group"
     Check UXValidationClose(), "Findings close failed"
   End Select
   Check UXGridPendingCount() = 1 And CurrentRevision(id) = before, "Read-only task saved or discarded pending cell"
   Check IsNull(DLookup("AssignmentRecommendationA", "Report", "ReportID=" & id)), "Read-only task committed invalid field"
  Next
  Check Not UXCurrentGroupTool("frmUXComparison", "Senior"), "Invalid cell allowed editing-tool navigation"
  Check Not UXGroupNavigate("frmUXComparison", 1), "Invalid cell allowed editing-context navigation"
  Check UXGridPendingCount() = 1, "Blocked navigation lost pending cell"
  Check UXGridDiscardPending(), "Explicit task-order discard failed"
  Check UXGridOpenReport(), "Task-order report did not open"
  Forms!frmEditor!txtMemberSSN = "123": before = CurrentRevision(id)
  Check CheckpointEditor(), "Invalid editor checkpoint failed"
  Check UXOpenNAVFITGroupExport(g), "Pending editor blocked saved group export dialog"
  Check CurrentProject.AllForms("frmEditor").IsLoaded And Forms!frmEditor!txtMemberSSN = "123", "Export dialog lost pending editor"
  Check CurrentRevision(id) = before And EditorHasPending(), "Export dialog saved or discarded editor"
  DoCmd.Close acForm, "frmNAVFITExchange", acSaveNo
  Check UXLeaveEditor("Discard"), "Explicit editor discard failed"
  Check UXBackToGroups(), "Task-order group exit failed"
 Next
 TestGroupTaskOrders = "PASS: " & n - start & " six-order pending-grid/editor, PDF, saved export and findings assertions": Exit Function
Failed:
 TestGroupTaskOrders = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function

Public Function TestUnifiedMembers() As String
 Dim g As Long, id As Long, other As Long, before As Long, snapshots As Long, i As Long, name As Variant
 On Error GoTo Failed
 Check UXNavigate("frmUXGroups"), "Unified setup navigation failed"
 DoCmd.OpenForm "frmUXNewGroup"
 Forms!frmUXNewGroup!txtTitle = "Unified empty Members"
 Forms!frmUXNewGroup!cboKind = "EVAL": Forms!frmUXNewGroup!cboPaygrade = "E5"
 Check UXCreateGroupUI(), "Create group UI failed"
 g = Forms!frmUXComparison!uxGroupContext
 Check g > 0 And Forms!frmUXComparison!uxTraitGrid.Form.Recordset.EOF, "New group does not have empty comparison table"
 Check Not Forms!frmUXComparison!uxOpenGroupReport.Enabled, "Empty roster allows opening a nonexistent selection"
 Check InStr(Forms!frmUXComparison!lblStatus.Caption, "No reports") > 0, "Empty Members guidance missing"
 Check Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded, "New group opened alternate workbench Members"
 For Each name In Array("uxMembers", "uxBulk", "uxPolicyReview", "uxOpenGroupReport", "uxValidation", "cmdBack", "uxCurrentSenior", "uxCurrentExport")
  Check Forms!frmUXComparison.Controls(name).Visible, "Missing unified member action: " & name
  Check Forms!frmUXComparison.Controls(name).Top >= Forms!frmUXComparison!uxTraitGrid.Top + Forms!frmUXComparison!uxTraitGrid.Height, "Action overlaps table: " & name
 Next
 before = DCount("*", "ReportRevision"): snapshots = DCount("*", "ComparisonSet")
 Check UXShowGroupMembers(g, True), "Empty refresh failed"
 Check DCount("*", "ReportRevision") = before And DCount("*", "ComparisonSet") = snapshots, "Empty open/refresh wrote history"
 id = CommitDraft(0, "EVAL", "Unified member", "Synthetic member", Null, Null, "")
 other = CreateSummaryGroup("Unrelated context", "EVAL", "E5")
 Check UXMemberAction("Members"), "Choose members failed"
 Check Not Forms!frmMajorWorkbench.Visible And Forms!frmMajorWorkbench!cboGroup = g, "Legacy adapter context is visible or incorrect"
 Forms!frmMajorWorkbench!cboGroup = other
 DoCmd.Close acForm, "frmUXMembers", acSaveNo
 Check UXMemberAction("Members"), "Reopening membership failed"
 Check Forms!frmMajorWorkbench!cboGroup = g, "Toolbar used stale hidden group"
 With Forms!frmUXMembers!lstReports
  For i = 0 To .ListCount - 1: .Selected(i) = (.ItemData(i) = id): Next
 End With
 Check UXPreviewGroupMembers(), "Membership preview failed"
 Check UXApplyGroupMembers(), "Membership apply failed"
 Check Forms!frmUXComparison!uxGroupContext = g And Not Forms!frmUXComparison!uxTraitGrid.Form.Recordset.EOF, "First member switched Members implementation"
 Check Not Forms!frmMajorWorkbench.Visible, "First member exposed workbench"
 before = CurrentRevision(id)
 Forms!frmUXComparison!uxTraitGrid.Form!AssignmentRecommendationA = String$(21, "X")
 Check UXGridQueue("AssignmentRecommendationA"), "Unified pending edit queue failed"
 Check Not UXMemberAction("Members"), "Toolbar abandoned invalid pending cell"
 Check UXGridPendingCount() = 1 And CurrentRevision(id) = before, "Blocked toolbar lost pending value or saved invalid revision"
 Check UXGridDiscardPending(), "Unified pending discard failed"
 Check UXMemberAction("Blocks"), "Unified shared-block editor failed"
 Check Forms!frmUXBulk!uxGroupContext = g, "Shared-block editor targeted wrong group"
 Check UXCloseBlocks(), "Shared-block return failed"
 Check UXCurrentGroupTool("frmUXComparison", "Senior"), "Unified reporting-senior action failed"
 Check Forms!frmUXGroupRSCA!txtGroupID = g, "Reporting-senior action targeted wrong group"
 DoCmd.Close acForm, "frmUXGroupRSCA", acSaveNo
 Check UXCurrentGroupTool("frmUXComparison", "Export"), "Unified export action failed"
 Check Forms!frmNAVFITExchange!cboGroup = g, "Export targeted wrong group"
 DoCmd.Close acForm, "frmNAVFITExchange", acSaveNo
 Check UXMemberAction("Validation"), "Unified group validation failed"
 Check Forms!frmUXValidation.OpenArgs = "G:" & g, "Validation targeted wrong group"
 DoCmd.Close acForm, "frmUXValidation", acSaveNo
 Check UXMemberAction("Policy"), "Unified policy review failed"
 Check Forms!frmUXGroupPolicy!groupID = g, "Policy targeted wrong group"
 DoCmd.Close acForm, "frmUXGroupPolicy", acSaveNo
 Check UXMemberAction("Open"), "Unified selected report failed"
 Check Forms!frmEditor!txtReportID = id, "Wrong selected report opened"
 Check UXFeedbackCloseEditor(), "Member return failed"
 Check UXGroupNavigate("frmUXComparison", 2), "Review tab failed"
 Check Forms!frmMajorWorkbench.Visible And Forms!frmMajorWorkbench!uxMode = "Review", "Review context lost"
 Check UXGroupNavigate("frmMajorWorkbench", 0), "Review return to Members failed"
 Check Not Forms!frmMajorWorkbench.Visible, "Review return exposed alternate Members"
 Check UXMemberAction("Members"), "Remove-last membership dialog failed"
 With Forms!frmUXMembers!lstReports
  For i = 0 To .ListCount - 1: .Selected(i) = False: Next
 End With
 Check UXPreviewGroupMembers() And UXApplyGroupMembers(), "Remove-last membership failed"
 Check Forms!frmUXComparison!uxTraitGrid.Form.Recordset.EOF, "Empty-after-removal changed view or retained row"
 Check Forms!frmUXComparison!uxMembers.Visible And Not Forms!frmMajorWorkbench.Visible, "Empty-after-removal lost actions"
 Check UXGroupAddReport("frmUXComparison"), "Empty-roster Add report failed"
 Forms!frmUXGroupReport!txtName = "Synthetic newly created member"
 Forms!frmUXGroupReport!txtOfficial = "SYNTHETIC, NEW": Forms!frmUXGroupReport!txtRank = "IT2"
 Check UXGroupReportCreate(), "Create report from empty Members failed"
 id = Forms!frmEditor!txtReportID
 Check PersistenceCount("GroupMember", "GroupID=" & g & " AND ReportID=" & id) = 1, "Created report lost group link"
 Check UXGroupCreatedReportClose(), "New report return failed"
 Check Forms!frmUXComparison!uxGroupContext = g And Not Forms!frmUXComparison!uxTraitGrid.Form.Recordset.EOF, "New report did not return to unified Members"
 Check UXBackToGroups(), "Unified empty exit failed"
 TestUnifiedMembers = "PASS: " & n & " total group exit and unified Members assertions": Exit Function
Failed:
 TestUnifiedMembers = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
