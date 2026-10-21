Attribute VB_Name = "UXOrganizationHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 6806, , message
    n = n + 1
End Sub
Public Function TestUXOrganization() As String
    Dim id As Long, g As Long, other As Long, bad As Long, rev As Long, fields As Object, stamp As String
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
