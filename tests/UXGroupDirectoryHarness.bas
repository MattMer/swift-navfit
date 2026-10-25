Attribute VB_Name = "UXGroupDirectoryHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 7069, , message
    checks = checks + 1
End Sub
Public Function TestUXGroupDirectory() As String
    Dim values As Object, readback As Object, key As Variant, memberID As Long, version As Long, rejected As Boolean, id As Long, groupID As Long, grade As String, kind As Variant, rate As String
    Dim baseline As String, defaults As String, reports As Long, members As Long, f As Access.Form, page As Long, i As Long, captions As Variant, screen As Variant, oldValue As Variant, width As Variant, w As DAO.Workspace, otherGroup As Long, operations As Long, revisions As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC GROUP DIRECTORY TEST", "Tester"
    Set values = RUADDecode(Null)
    Check values.Count = 35, "RUAD must retain all 35 columns"
    For Each key In values.Keys: values(key) = "Synthetic " & key & " & apostrophe ' and <tag>": Next
    values("UIC") = "00123": values("RUIC") = "00001": values("PRD") = "30-Jun-31": values("PHONE") = Null: values("ADDRESS") = ""
    memberID = SaveDirectoryMember(0, 0, "Synthetic RUAD member", "SYNTHETIC, RUAD", "IT2", "", "Manual", values)
    Set readback = RUADDecode(DLookup("RUADData", "CommandMember", "MemberID=" & memberID))
    For Each key In values.Keys
        If key = "NAME" Then
            Check readback(key) = "Synthetic RUAD member", "NAME differs from core member name"
        Else
            Check SameFieldValue(readback(key), values(key)), "RUAD round trip failed: " & key
        End If
    Next
    version = DLookup("VersionCounter", "CommandMember", "MemberID=" & memberID)
    SaveDirectoryMember memberID, version, "Renamed synthetic RUAD", "SYNTHETIC, RUAD", "IT2", ""
    Set readback = RUADDecode(DLookup("RUADData", "CommandMember", "MemberID=" & memberID))
    Check readback("UIC") = "00123" And readback("NAME") = "Renamed synthetic RUAD", "Basic profile save erased RUAD or left NAME stale"
    oldValue = DLookup("RUADData", "CommandMember", "MemberID=" & memberID)
    On Error Resume Next
    SaveDirectoryMember memberID, version, "Stale", "", "IT2", "", "Manual", values
    rejected = (Err.Number <> 0): Err.Clear: On Error GoTo Failed
    Check rejected And DLookup("RUADData", "CommandMember", "MemberID=" & memberID) = oldValue, "Stale member edit altered data"
    values.Add "Unknown column", "Bad"
    version = DLookup("VersionCounter", "CommandMember", "MemberID=" & memberID)
    On Error Resume Next
    SaveDirectoryMember memberID, version, "Rejected", "", "IT2", "", "Manual", values
    rejected = (Err.Number <> 0): Err.Clear: On Error GoTo Failed
    Check rejected And DLookup("VersionCounter", "CommandMember", "MemberID=" & memberID) = version, "Unknown RUAD field did not roll back"
    values.Remove "Unknown column"
    DoCmd.OpenForm "frmUXDirectory"
    Forms!frmUXDirectory!lstMembers = memberID
    Check UXDirectoryEdit(False), "Profile editor did not open"
    Set f = Forms!frmUXMemberEditor
    Check f!ruad5 = "00123", "Profile did not load UIC"
    f!ruad5 = "00009"
    For page = 0 To 4
        Check UXRUADPage(page), "Profile page failed"
        Check f!ruad5 = "00009", "Page navigation lost pending field"
    Next
    Check UXDirectorySave(), "Profile save failed"
    Set readback = RUADDecode(DLookup("RUADData", "CommandMember", "MemberID=" & memberID))
    Check readback("UIC") = "00009", "Profile save did not persist edited UIC"
    DoCmd.Close acForm, "frmUXDirectory", acSaveNo
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        Select Case CStr(kind)
            Case "EVAL": grade = "E5": rate = "IT2"
            Case "CHIEFEVAL": grade = "E7": rate = "ITC"
            Case Else: grade = "O3": rate = "LT"
        End Select
        groupID = CreateSummaryGroup("Synthetic group " & kind, CStr(kind), grade)
        SetFieldDefault "Group", groupID, "UnitUIC", "00123"
        SetFieldDefault "Group", groupID, "CommandEmployment", "Synthetic command achievements"
        baseline = GroupStamp(groupID): defaults = UXDefaultsStamp(groupID)
        id = CreateGroupDraft(groupID, baseline, defaults, 0, 0, "Synthetic " & kind, "SYNTHETIC, MEMBER", rate, "", Null, Null)
        Check DLookup("FormKind", "Report", "ReportID=" & id) = kind, "Group report type incorrect"
        Check DCount("*", "GroupMember", "GroupID=" & groupID & " AND ReportID=" & id) = 1, "New report not added to group"
        Check DLookup("UnitUIC", "Report", "ReportID=" & id) = "00123", "Group default missing"
        Check DLookup("CommandEmployment", "Report", "ReportID=" & id) = "Synthetic command achievements", "Block 28 default missing"
        Check DCount("*", "MemberReport", "ReportID=" & id) = 1, "New member directory link missing"
        Check DCount("*", "ReportRevision", "ReportID=" & id) = 1, "Creation should make one revision"
        reports = DCount("*", "Report"): members = DCount("*", "CommandMember")
        On Error Resume Next
        CreateGroupDraft groupID, baseline, defaults, 0, 0, "Rejected stale", "STALE", rate, "", Null, Null
        rejected = (Err.Number <> 0): Err.Clear: On Error GoTo Failed
        Check rejected And reports = DCount("*", "Report") And members = DCount("*", "CommandMember"), "Stale group creation did not roll back"
    Next
    groupID = CreateSummaryGroup("Atomic caller rollback", "EVAL", "E5")
    reports = DCount("*", "Report"): members = DCount("*", "CommandMember"): operations = DCount("*", "OperationJournal")
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    id = CreateGroupDraft(groupID, GroupStamp(groupID), UXDefaultsStamp(groupID), 0, 0, "Rolled-back member", "ROLLED BACK", "IT2", "", Null, Null)
    w.Rollback
    Check DCount("*", "Report") = reports And DCount("*", "CommandMember") = members And DCount("*", "OperationJournal") = operations And DCount("*", "GroupMember", "GroupID=" & groupID) = 0, "Report/directory/group/journal did not roll back atomically"
    groupID = CreateSummaryGroup("Existing directory creation", "EVAL", "E5")
    version = DLookup("VersionCounter", "CommandMember", "MemberID=" & memberID)
    id = CreateGroupDraft(groupID, GroupStamp(groupID), UXDefaultsStamp(groupID), memberID, version, "Renamed synthetic RUAD", "SYNTHETIC, RUAD", "IT2", "", #1/1/2026#, #9/30/2026#)
    Check DLookup("MemberID", "MemberReport", "ReportID=" & id) = memberID, "Existing directory selection lost"
    reports = DCount("*", "Report")
    On Error Resume Next
    CreateGroupDraft groupID, GroupStamp(groupID), UXDefaultsStamp(groupID), memberID, version, "Duplicate", "SYNTHETIC, RUAD", "IT2", "", #1/1/2026#, #9/30/2026#
    rejected = (Err.Number <> 0): Err.Clear: On Error GoTo Failed
    Check rejected And DCount("*", "Report") = reports, "Duplicate type and period accepted"
    DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = groupID: Forms!frmMajorWorkbench!uxGroupContext = groupID
    Check UXWorkbenchMode("Members"), "Group members mode failed"
    captions = Array("Members", "Defaults", "Review", "", "History")
    Check UXGroupAddReport("frmMajorWorkbench"), "Add report dialog failed"
    DoCmd.Close acForm, "frmUXGroupReport", acSaveNo
    Check DCount("*", "Report") = reports, "Cancel created an empty report"
    Check UXGroupNavigate("frmMajorWorkbench", 0), "Members navigation failed"
    Check Forms!frmUXComparison!uxGroupContext = groupID, "Members context lost"
    operations = DCount("*", "ComparisonSet")
    Check UXGroupNavigate("frmUXComparison", 3), "Comparison navigation failed"
    Check DCount("*", "ComparisonSet") = operations And Forms!frmUXComparison!uxGroupTab0.FontBold And Not UXHasControl(Forms!frmUXComparison, "uxGroupTab3"), "Comparison navigation recreated a set or left members"
    Check UXGroupNavigate("frmUXComparison", 0), "Members tab switch failed"
    Check DCount("*", "ComparisonSet") = operations And Forms!frmUXComparison!uxGroupTab0.FontBold, "Tab switch recreated or failed to select members"
    otherGroup = CreateSummaryGroup("Other group context", "FITREP", "O3")
    Forms!frmMajorWorkbench!cboGroup = otherGroup
    UXGridViewport
    Check Forms!frmUXComparison!uxGroupContext = groupID, "Resize retargeted an open comparison to another group"
    Forms!frmMajorWorkbench!cboGroup = groupID
    For i = 0 To 4: If i <> 3 Then Check Forms!frmUXComparison.Controls("uxGroupTab" & i).Caption = captions(i), "Comparison tab caption wrong"
    Next
    Check UXGroupNavigate("frmUXComparison", 1), "Defaults navigation failed"
    For i = 0 To 4: If i <> 3 Then Check Forms!frmUXBulk.Controls("uxGroupTab" & i).Caption = captions(i) And Forms!frmUXBulk.Controls("uxGroupTab" & i).Visible, "Defaults tab caption/visibility wrong"
    Next
    Check Forms!frmUXBulk!uxGroupContext = groupID And Forms!frmUXBulk!uxAddGroupReport.Visible, "Defaults group context/Add report missing"
    revisions = DCount("*", "ReportRevision")
    Forms!frmUXBulk!uxValueUnitUIC = "00001": Call UXConfirmBlockChange("UnitUIC", vbNo)
    Check Nz(Forms!frmUXBulk!uxValueUnitUIC, "") <> "00001" And DCount("*", "ReportRevision") = revisions, "Declined default change was applied"
    UXCloseBlocks
    Check UXGroupNavigate("frmMajorWorkbench", 2), "Review navigation failed"
    Check Forms!frmMajorWorkbench!uxMode = "Review", "Review mode wrong"
    Check UXGroupNavigate("frmMajorWorkbench", 4), "History navigation failed"
    Check Forms!frmMajorWorkbench!uxMode = "History", "History mode wrong"
    For i = 0 To 4: If i <> 3 Then Check Forms!frmMajorWorkbench.Controls("uxGroupTab" & i).Caption = captions(i), "Main group tab caption wrong"
    Next
    Check UXGroupAddReport("frmMajorWorkbench"), "Create dialog failed"
    With Forms!frmUXGroupReport
        !txtName = "Synthetic group UI creation": !txtOfficial = "SYNTHETIC, UI": !txtRank = "IT2"
    End With
    Check UXGroupReportCreate(), "Native create action failed"
    Check Forms!frmEditor!cmdClose.Caption = "Back to group members", "Creation lost its return context"
    id = Forms!frmEditor!txtReportID
    Check DCount("*", "GroupMember", "GroupID=" & groupID & " AND ReportID=" & id) = 1, "Native-created report missing from group"
    Check UXGroupCreatedReportClose(), "Return from created report failed"
    Check CurrentProject.AllForms("frmUXComparison").IsLoaded And Forms!frmUXComparison!uxGroupContext = groupID, "Created report did not return to group members"
    DoCmd.OpenForm "frmHome"
    For Each width In Array(10000, 14000, 22000)
        Check UXResizeWorkspace("frmHome", CLng(width), 9500), "Overview viewport failed"
        Set f = Forms!frmHome
        Check f!uxContributor.Top + f!uxContributor.Height <= f!cmdMyWork.Top, "Contributor overlaps My Work"
        Check f!uxContributorStatus.Top + f!uxContributorStatus.Height <= f!cmdMyWork.Top, "Contributor status overlaps My Work"
        Check f!uxNoteRecovery.Top >= f!cmdRecovery.Top + f!cmdRecovery.Height And f!uxNoteRecovery.Top + f!uxNoteRecovery.Height < f!uxRecentCaption.Top, "Unposted notes overlaps recovery or recent reports"
    Next
    TestUXGroupDirectory = "PASS: " & checks & " group/RUAD/Overview checks": Exit Function
Failed:
    TestUXGroupDirectory = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
