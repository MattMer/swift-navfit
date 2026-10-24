Attribute VB_Name = "UXGroupLifecycleHarness"
Option Explicit
Private n As Long
Private workflowLookups As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 5600, , message
    n = n + 1
End Sub
Public Sub SummaryListCountWorkflowLookup()
    workflowLookups = workflowLookups + 1
End Sub
Public Function SummaryListLookupCount() As Long
    SummaryListLookupCount = workflowLookups
End Function
Public Function SeedSummaryListBenchmark(ByVal groups As Long, ByVal members As Long) As String
    Dim i As Long, j As Long, group As Long, report As Long, finish As Variant, states As Variant
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC SUMMARY LIST", "Synthetic list tester"
    states = Array("Initialized", "In-Work", "Ready for Signature")
    For i = 1 To groups
        group = CreateSummaryGroup("SYNTHETIC " & Format$(i, "0000"), "EVAL", "E5")
        For j = 1 To members
            finish = #9/30/2026#
            If i Mod 3 = 1 And j = members Then finish = Null
            If i Mod 3 = 2 And j = members Then finish = #8/31/2026#
            report = CommitDraft(0, "EVAL", "SYNTHETIC " & i & "/" & j, "SYNTHETIC MEMBER " & i & "/" & j, #1/1/2026#, finish, "")
            AddGroupMember group, report, "E5"
        Next
        If i Mod 3 > 0 Then SetGroupWorkflow group, GroupStamp(group), CStr(states(i Mod 3))
        If i Mod 5 = 0 Then ChangeSummaryGroup group, GroupStamp(group), "Archive"
    Next
    SetFieldDefault "Command", 1, "PeriodEnd", "2026-09-30"
    group = CreateSummaryGroup("SYNTHETIC empty inherited", "EVAL", "E5")
    group = CreateSummaryGroup("SYNTHETIC empty override", "EVAL", "E5"): SetFieldDefault "Group", group, "PeriodEnd", "2026-08-31"
    group = CreateSummaryGroup("SYNTHETIC empty blank", "EVAL", "E5"): SetFieldDefault "Group", group, "PeriodEnd", ""
    group = CreateSummaryGroup("SYNTHETIC empty null", "EVAL", "E5"): SetFieldDefault "Group", group, "PeriodEnd", Null
    group = CreateSummaryGroup("SYNTHETIC missing lifecycle", "EVAL", "E5")
    CurrentDb.Execute "DELETE FROM GroupLifecycle WHERE GroupID=" & group, dbFailOnError
    'Read-only legacy-state equivalence, not an RSCA completion/posting fixture.
    group = CreateSummaryGroup("SYNTHETIC retained completed", "EVAL", "E5")
    CurrentDb.Execute "DELETE FROM GroupLifecycle WHERE GroupID=" & group, dbFailOnError
    CurrentDb.Execute "INSERT INTO GroupLifecycle (GroupID,WorkflowState) VALUES (" & group & ",'Completed')", dbFailOnError
    group = CreateSummaryGroup("SYNTHETIC same title", "EVAL", "E5")
    group = CreateSummaryGroup("SYNTHETIC same title", "EVAL", "E5")
    SeedSummaryListBenchmark = "PASS: " & groups & " populated groups / " & groups * members & " reports plus eight read-boundary groups": Exit Function
Failed:
    Dim message As String
    message = Err.Number & " " & Err.Description
    SeedSummaryListBenchmark = "FAIL: summary-list fixture: " & message
End Function
Public Function SummaryListRead(ByVal filter As String) As String
    Dim row As Long, col As Long, value As Variant, text As String, f As Access.Form
    If Not CurrentProject.AllForms("frmUXGroups").IsLoaded Then DoCmd.OpenForm "frmUXGroups"
    Set f = Forms!frmUXGroups: f!uxArchiveFilter = filter
    workflowLookups = 0
    If Not UXGroupList() Then Err.Raise 5, , "Summary list did not refresh"
    For row = 0 To f!lstGroups.ListCount - 1
        For col = 0 To 7
            value = f!lstGroups.Column(col, row)
            If IsNull(value) Then text = text & "N;" Else text = text & Len(CStr(value)) & ":" & CStr(value) & ";"
        Next
        text = text & vbLf
    Next
    SummaryListRead = text
End Function
Public Function SummaryListDataCounts() As String
    Dim table As Variant
    For Each table In Array("Report", "ReportRevision", "RevisionValue", "SummaryGroup", "GroupMember", "GroupLifecycle", "FieldDefault", "OperationJournal", "GroupSnapshot", "ComparisonSet")
        SummaryListDataCounts = SummaryListDataCounts & table & "=" & DCount("*", CStr(table)) & ";"
    Next
End Function
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
    Check UXOpenGroup(True), "Group context setup failed": UXWorkbenchMode "Members"
    proposal = ProposeMeetingField(meeting, id, "Trait33", "4.0")
    Check UXStartGroupBulk(), "Bulk setup failed"
    Forms!frmUXBulk!uxValueShipStation = "CHANGED": Check Not UXConfirmBlockChange("ShipStation", vbNo), "Cancelled shared edit applied"
    Check Not UXGroupBulkPending(), "Shared editor retained staging state"
    stage = "rename dialog"
    DoCmd.OpenForm "frmUXRenameGroup", , , , , , CStr(g)
    Forms!frmUXRenameGroup!txtName = "Renamed O'Neil group"
    Check UXSaveGroupName(), "Rename save failed"
    Check DLookup("Title", "SummaryGroup", "GroupID=" & g) = "Renamed O'Neil group", "Rename not stored"
    Check Forms!frmUXComparison!lblHeading.Caption = "Renamed O'Neil group", "Open workspace title stale"
    Check Rejected(g, stamp, "Archive"), "Stale lifecycle write accepted"
    Check RejectMeeting(proposal), "Comparison survived renamed baseline"
    Check Not UXGroupBulkPending(), "Rename created pending shared changes"
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
    Check UXOpenGroup(True), "Archived context setup failed": UXWorkbenchMode "Members"
    Check Not Forms!frmUXComparison!uxBulk.Enabled And Not Forms!frmUXComparison!uxMembers.Enabled, "Archived workspace editable"
    Check InStr(Forms!frmUXComparison!lblStatus.Caption, "Archived group") > 0, "Archived reason missing"
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
    Check Forms!frmUXComparison!uxBulk.Enabled And Forms!frmUXComparison!uxMembers.Enabled, "Recovered workspace still locked"
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
