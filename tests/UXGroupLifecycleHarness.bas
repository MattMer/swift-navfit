Attribute VB_Name = "UXGroupLifecycleHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 5600, , message
    n = n + 1
End Sub
Private Function Rejected(ByVal id As Long, ByVal stamp As String, ByVal action As String, Optional ByVal title As String = "") As Boolean
    On Error GoTo Failed
    ChangeSummaryGroup id, stamp, action, title: Exit Function
Failed:
    Rejected = True
End Function
Private Function RejectDefault(ByVal id As Long) As Boolean
    On Error GoTo Failed
    SetFieldDefault "Group", id, "UnitUIC", "99999": Exit Function
Failed:
    RejectDefault = True
End Function
Private Function RejectMeeting(ByVal proposal As Long) As Boolean
    On Error GoTo Failed
    ApplyMeetingProposals CStr(proposal): Exit Function
Failed:
    RejectMeeting = True
End Function
Public Function TestUXGroupLifecycle() As String
    Dim g As Long, id As Long, rev As Long, stamp As String, snap As Long, meeting As Long, version As Long, proposal As Long, f As Access.Form, stage As String
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC LIFECYCLE", "Tester"
    id = CommitDraft(0, "EVAL", "Test report", "Test member", #1/1/2026#, #9/30/2026#, "", "Test", "")
    rev = CurrentRevision(id): g = CreateSummaryGroup("Test group", "EVAL", "E5"): AddGroupMember g, id, "E5"
    SetFieldDefault "Group", g, "UnitUIC", "12345"
    snap = FreezeSummaryGroup(g, "Test"): meeting = CreateComparisonSet(snap, "Test meeting")
    stamp = GroupStamp(g): stage = "rename validation"
    Check Rejected(g, stamp, "Rename", ""), "Blank name accepted"
    Check Rejected(g, stamp, "Rename", String$(121, "X")), "Overlong name accepted"
    Check Rejected(g, stamp, "Unknown"), "Unknown action accepted"
    Check GroupStamp(g) = stamp, "Failed action changed group"
    DoCmd.OpenForm "frmUXGroups": Set f = Forms!frmUXGroups
    Check f!lstGroups.AfterUpdate = "=UXGroupSelection()", "Keyboard selection event missing"
    f!lstGroups = g: UXGroupSelection
    Check f!uxRename.Enabled And f!uxArchive.Enabled And Not f!uxRecover.Enabled, "Active actions incorrect"
    Check UXOpenGroup(), "Group open failed"
    proposal = ProposeMeetingField(meeting, id, "Trait33", "4.0")
    Check UXStartGroupBulk(), "Bulk setup failed"
    Forms!frmUXBulk!uxValueShipStation = "CHANGED": UXBulkChanged "ShipStation": Forms!frmUXBulk!txtReason = "Test stale rename"
    Check UXPreviewBulk(), "Bulk preview setup failed"
    stage = "rename dialog"
    DoCmd.OpenForm "frmUXRenameGroup", , , , , , CStr(g)
    Forms!frmUXRenameGroup!txtName = "Renamed O'Neil group"
    Check UXSaveGroupName(), "Rename save failed"
    Check DLookup("Title", "SummaryGroup", "GroupID=" & g) = "Renamed O'Neil group", "Rename not stored"
    Check Forms!frmMajorWorkbench!uxGroupTitle.Caption = "Renamed O'Neil group", "Open workspace title stale"
    Check Rejected(g, stamp, "Archive"), "Stale lifecycle write accepted"
    Check RejectMeeting(proposal), "Comparison survived renamed baseline"
    Check Not UXApplyBulk(), "Bulk preview survived renamed baseline"
    UXCloseBlocks
    Check CurrentRevision(id) = rev And DCount("*", "GroupMember", "GroupID=" & g) = 1, "Rename changed report or members"
    stage = "archive"
    f!lstGroups = g: f!uxArchive.SetFocus
    Check UXGroupAction("Archive"), "Archive action failed"
    Check DLookup("IsArchived", "SummaryGroup", "GroupID=" & g), "Archive not stored"
    Check Nz(f!lstGroups, 0) = 0 And Not f!uxRecover.Enabled, "Hidden archived group remains selected"
    f!uxArchiveFilter = "Archived": Check UXGroupList(), "Archive filter failed"
    f!lstGroups = g: Check UXGroupSelection(), "Archive selection failed"
    Check f!uxRecover.Enabled And Not f!uxRename.Enabled And Not f!uxArchive.Enabled, "Archived actions incorrect"
    Check f!lstGroups.Column(5) = "Archived" And CLng(f!lstGroups.Column(4)) = 1, "Archive status/count wrong"
    Check UXOpenGroup(), "Archived inspection failed"
    Check Not Forms!frmMajorWorkbench!uxBulk.Enabled And Not Forms!frmMajorWorkbench!uxMembers.Enabled, "Archived workspace editable"
    Check InStr(Forms!frmMajorWorkbench!lblStatus.Caption, "Archived group") > 0, "Archived reason missing"
    Check Not UXStartGroupBulk(), "Archived bulk opened"
    Check RejectDefault(g), "Archived defaults writable"
    Check Rejected(g, GroupStamp(g), "Rename", "Forbidden"), "Archived rename accepted"
    Check DCount("*", "GroupSnapshot", "SnapshotID=" & snap) = 1 And DCount("*", "ComparisonSet", "SetID=" & meeting) = 1, "History removed"
    Check UXGroupDefaultValue(g, "UnitUIC") = "12345", "Default removed"
    Check CurrentRevision(id) = rev And Not DLookup("IsArchived", "Report", "ReportID=" & id), "Group archived member report"
    stage = "recover"
    f!uxRecover.SetFocus: Check UXGroupAction("Recover"), "Recover action failed"
    Check Not DLookup("IsArchived", "SummaryGroup", "GroupID=" & g), "Recover not stored"
    f!uxArchiveFilter = "All": UXGroupList: f!lstGroups = g: UXGroupSelection
    Check f!lstGroups.Column(5) = "Active" And f!uxArchive.Enabled, "Recovered list status wrong"
    Check Forms!frmMajorWorkbench!uxBulk.Enabled And Forms!frmMajorWorkbench!uxMembers.Enabled, "Recovered workspace still locked"
    Check CurrentRevision(id) = rev And DCount("*", "GroupMember", "GroupID=" & g) = 1, "Recovery changed reports/members"
    Check Rejected(g, GroupStamp(g), "Recover"), "Repeated recovery accepted"
    Check DCount("*", "OperationJournal", "OperationKind IN ('Rename summary group','Archive summary group','Recover summary group')") = 3, "Lifecycle audit missing"
    stage = "stale rename buffer"
    DoCmd.OpenForm "frmUXRenameGroup", , , , , , CStr(g)
    Forms!frmUXRenameGroup!txtName = "Pending name"
    ChangeSummaryGroup g, GroupStamp(g), "Rename", "Concurrent name"
    Check Not UXSaveGroupName(), "Stale rename saved"
    Check Forms!frmUXRenameGroup!txtName = "Pending name", "Failed rename lost input"
    DoCmd.Close acForm, "frmUXRenameGroup", acSaveNo
    TestUXGroupLifecycle = "PASS: " & n & " group lifecycle assertions": Exit Function
Failed:
    TestUXGroupLifecycle = "FAIL: " & stage & ": " & Err.Number & " " & Err.Description
End Function
