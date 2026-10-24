Attribute VB_Name = "UXOrganizationHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6806, , message
    n = n + 1
End Sub
Public Function TestMoveDisplayFailure() As String
    Dim report As Long, source As Long, destination As Long, fields As Object, revision As Long
    On Error GoTo Failed
    n = 0: ConfigureWorkspace "SYNTHETIC MOVE DISPLAY FAILURE", "Synthetic mover"
    report = CommitDraft(0, "EVAL", "Move display fixture", "SYNTHETIC MEMBER", Null, Null, "")
    Set fields = CreateObject("Scripting.Dictionary"): fields.Add "RankRate", "IT2"
    CommitSavedFields report, CurrentRevision(report), fields, "Synthetic worn rate": revision = CurrentRevision(report)
    source = CreateSummaryGroup("Original group", "EVAL", "E5"): destination = CreateSummaryGroup("Destination group", "EVAL", "E5")
    AddGroupMember source, report, "E5"
    DoCmd.OpenForm "frmUXMoveGroup", , , , , , CStr(report)
    Forms!frmUXMoveGroup!cboGroup = destination
    Check UXPreviewMoveGroup(), "Display-fault move preview failed"
    Check Not UXApplyMoveGroup(), "Injected move display failure was accepted"
    Check DCount("*", "GroupMember", "ReportID=" & report & " AND GroupID=" & destination) = 1 And DCount("*", "GroupMember", "ReportID=" & report & " AND GroupID=" & source) = 0, "Display failure rolled back the completed move"
    Check CurrentRevision(report) = revision, "Display failure changed official report values"
    Check InStr(Forms!frmUXMoveGroup!lblStatus.Caption, "Move completed") > 0 And InStr(Forms!frmUXMoveGroup!lblStatus.Caption, "SYNTHETIC") > 0, "Committed move/display failure not disclosed"
    Check Not Forms!frmUXMoveGroup!cmdApply.Enabled And Not UXApplyMoveGroup(), "Display failure retained an actionable move"
    Check DCount("*", "OperationJournal", "OperationKind='Move report to summary group' AND OperationState='Completed'") = 1, "Repeated Apply duplicated move history"
    TestMoveDisplayFailure = "PASS: " & n & " committed-move/display-failure assertions": Exit Function
Failed:
    TestMoveDisplayFailure = "FAIL: move display fault: " & Err.Number & " " & Err.Description
End Function
Public Function TestUXOrganization() As String
    Dim id As Long, g As Long, other As Long, bad As Long, rev As Long, fields As Object, stamp As String, owner As Long, second As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC ORGANIZATION", "Tester"
    id = CommitDraft(0, "EVAL", "Group move", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test", "")
    Set fields = CreateObject("Scripting.Dictionary"): fields.Add "RankRate", "IT2"
    CommitSavedFields id, CurrentRevision(id), fields, "Set worn rate"
    g = CreateSummaryGroup("Source group", "EVAL", "E5"): other = CreateSummaryGroup("Destination group", "EVAL", "E5")
    bad = CreateSummaryGroup("Wrong grade", "EVAL", "E6")
    AddGroupMember g, id, "E5": rev = CurrentRevision(id)
    DoCmd.OpenForm "frmLibrary": RefreshFolders
    Check Not UXHasControl(Forms!frmLibrary, "cboDestination"), "Obsolete destination folder control remains"
    Forms!frmLibrary!uxFolderFilter = g
    Check UXFolderFilterChanged(), "Group filter failed"
    Check Forms!frmLibrary!lstReports.ListCount = 2, "Group filter report missing"
    Forms!frmLibrary!lstReports = id
    Check UXChooseReportGroup(), "Move popup did not open"
    Check Not UXApplyMoveGroup(), "Move bypassed preview"
    Forms!frmUXMoveGroup!cboGroup = bad
    Check Not UXPreviewMoveGroup(), "Incompatible paygrade allowed"
    Forms!frmUXMoveGroup!cboGroup = other
    Check UXPreviewMoveGroup(), "Move preview failed"
    Check InStr(Forms!frmUXMoveGroup!txtPreview, "Source group") > 0 And InStr(Forms!frmUXMoveGroup!txtPreview, "Destination group") > 0, "Move before/after labels absent"
    DoCmd.Close acForm, "frmUXMoveGroup", acSaveNo
    DoCmd.OpenForm "frmUXMoveGroup", , , , , , CStr(id)
    Forms!frmUXMoveGroup!cboGroup = other
    Check Not UXApplyMoveGroup() And Not Forms!frmUXMoveGroup!cmdApply.Enabled, "Reopened move reused a previous preview"
    Check UXPreviewMoveGroup(), "Reopened move could not preview again"
    owner = ActiveContributorID(): second = AddContributor("Second synthetic mover"): SwitchContributor second
    Check Not UXApplyMoveGroup(), "Another contributor applied the reviewed move"
    Check DCount("*", "GroupMember", "ReportID=" & id & " AND GroupID=" & g) = 1, "Rejected move ownership changed membership"
    SwitchContributor owner
    Check Not UXApplyMoveGroup(), "Rejected move ownership retained actionable preview"
    Check UXPreviewMoveGroup(), "Original contributor could not preview again"
    Check UXApplyMoveGroup(), "Move apply failed"
    Check DCount("*", "GroupMember", "ReportID=" & id) = 1 And DCount("*", "GroupMember", "ReportID=" & id & " AND GroupID=" & other) = 1, "Move did not leave previous group"
    Check CurrentRevision(id) = rev, "Membership move changed report values"
    Forms!frmUXMoveGroup!cboGroup = 0
    Check UXPreviewMoveGroup(), "Ungroup preview failed"
    fields("RankRate") = "IT2": fields.Add "OfficialComments", "Concurrent edit"
    CommitSavedFields id, CurrentRevision(id), fields, "Change after preview"
    Check Not UXApplyMoveGroup(), "Stale move accepted"
    Check DCount("*", "GroupMember", "ReportID=" & id & " AND GroupID=" & other) = 1, "Rejected move changed membership"
    Check UXPreviewMoveGroup() And UXApplyMoveGroup(), "Fresh ungroup move failed"
    Forms!frmLibrary!uxFolderFilter = -1: UXFolderFilterChanged
    Check Forms!frmLibrary!lstReports.ListCount = 2, "Ungrouped filter failed"
    DoCmd.Close acForm, "frmUXMoveGroup", acSaveNo
    Check GroupWorkflow(g) = "Initialized", "Initial group status wrong"
    stamp = GroupStamp(g)
    SetGroupWorkflow g, stamp, "In-Work"
    Check GroupWorkflow(g) = "In-Work", "Group status not persisted"
    On Error Resume Next
    SetGroupWorkflow g, stamp, "Completed"
    bad = Err.Number: Err.Clear: On Error GoTo Failed
    Check bad <> 0 And GroupWorkflow(g) = "In-Work", "Stale group status accepted"
    Check DLookup("WorkflowState", "Report", "ReportID=" & id) = "Draft", "Group status changed member report state"
    DoCmd.OpenForm "frmUXGroups": Forms!frmUXGroups!lstGroups = g: UXGroupSelection
    Forms!frmUXGroups!uxGroupState = "Ready for Signature"
    Check UXSetGroupWorkflow(), "Group-list status action failed"
    Check Forms!frmUXGroups!lstGroups.ColumnCount = 8 And GroupWorkflow(g) = "Ready for Signature" And Nz(Forms!frmUXGroups!lstGroups.Column(6), "") = "Ready for Signature", "Group-list status not shown"
    TestUXOrganization = "PASS: " & n & " group move/filter/status assertions": Exit Function
Failed:
    TestUXOrganization = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
