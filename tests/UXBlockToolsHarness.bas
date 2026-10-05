Attribute VB_Name = "UXBlockToolsHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 5500, , message
    checks = checks + 1
End Sub
Public Function TestUXBlockTools() As String
    Dim id As Long, otherID As Long, g As Long, rev As Long, otherRev As Long, values As Object, f As Access.Form, rows As Access.Form, stage As String, key As Variant, p As Long, snapshot As Long, meeting As Long, kind As Variant, cache As String, query As DAO.QueryDef, db As DAO.Database
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC BLOCK TEST", "Tester"
    id = CommitDraft(0, "EVAL", "First", "First", #1/1/2026#, #9/30/2026#, "", "Test", "")
    otherID = CommitDraft(0, "EVAL", "Second", "Second", #1/1/2026#, #9/30/2026#, "", "Test", "")
    Set values = CreateObject("Scripting.Dictionary"): values("ShipStation") = "ALPHA"
    For Each key In Array("Trait33", "Trait34", "Trait35", "Trait36", "Trait37", "Trait38", "Trait39"): values(key) = "3.0": Next
    CommitSavedFields id, CurrentRevision(id), values, "Traits"
    values("ShipStation") = "BRAVO": CommitSavedFields otherID, CurrentRevision(otherID), values, "Traits"
    g = CreateSummaryGroup("Block test", "EVAL", "E5"): AddGroupMember g, id, "E5": AddGroupMember g, otherID, "E5"
    DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = g
    stage = "shared blocks"
    Check UXStartGroupBulk(), "Bulk open failed": Set f = Forms!frmUXBulk
    Check f!uxBlocks.Visible, "Missing native shared surface"
    Check f!uxBlocks.Form!stateShipStation.Caption = "Mixed values - unchanged", "Mixed value ambiguous"
    Check f!uxBlocks.Form!valueUnitUIC.Visible, "Section fields missing"
    Check UXBulkSection(1), "Senior section failed"
    Check f!uxBlocks.Form!valueReportingSeniorName.Visible, "Senior fields unavailable"
    Check UXBulkSection(0), "Member section failed"
    f!uxBlocks.Form!valueShipStation = "CHARLIE": Check UXBlockEdit("ShipStation"), "Surface edit not bridged"
    f!txtReason = "Shared command": Check UXPreviewBulk(), "Bulk preview failed"
    Check f!lstPreview.ListCount = 2 And Not f!uxBlocks.Visible, "Preview scope wrong"
    rev = CurrentRevision(id): Check UXApplyBulk(), "Bulk apply failed"
    Check CurrentRevision(id) <> rev And DLookup("ShipStation", "Report", "ReportID=" & otherID) = "CHARLIE", "Atomic reports not saved"
    UXCloseBlocks
    stage = "defaults"
    Check UXOpenGroupDefaults(), "Defaults open failed": Set f = Forms!frmUXBulk
    f!uxBlocks.Form!valueUnitUIC = "54321": UXBlockEdit "UnitUIC"
    f!uxBlocks.Form!valuePeriodStart = "2/1/2026": UXBlockEdit "PeriodStart": f!txtReason = "Default unit and period"
    rev = CurrentRevision(id): Check UXPreviewBulk(), "Defaults preview failed"
    Check f!lstPreview.ListCount = 2, "Defaults preview has incorrect scope"
    Check UXApplyBulk(), "Defaults apply failed"
    Check CurrentRevision(id) = rev, "Defaults changed saved reports"
    Check UXGroupDefaultValue(g, "UnitUIC") = "54321", "Default not stored"
    Check UXGroupDefaultValue(g, "PeriodStart") = "2026-02-01", "Date default not canonicalized"
    f!lstPreview = "UnitUIC": Check UXBulkDetail(), "Defaults result detail unavailable"
    Check InStr(f!txtFullValues, "AFTER: 54321") > 0 And InStr(f!txtFullValues, "BEFORE: (no value)") > 0, "Default Results changed the preview snapshot"
    Check UXBlockClear("UnitUIC"), "Clear not staged": Check UXPreviewBulk(), "Null preview failed"
    SetFieldDefault "Group", g, "UnitUIC", "11111"
    Check Not UXApplyBulk(), "Concurrent default change accepted"
    UXCloseBlocks
    stage = "editable comparison"
    Check UXOpenComparison(), "Comparison open failed": Set f = Forms!frmUXComparison: Set rows = f!uxTraitGrid.Form
    Check rows.Recordset.RecordCount = 2, "Editable rows missing"
    Check rows!ITA = "3.00" And InStr(f!uxGridAverage.Caption, "3.00") > 0, "Initial saved-baseline averages missing"
    SetMeetingOrder f!txtMeeting, CStr(otherID) & "," & CStr(id)
    Check UXGridLoad(), "Ordered grid reload failed"
    Set rows = f!uxTraitGrid.Form
    Check rows!ReportID = otherID, "Comparison display order lost"
    rev = CurrentRevision(id): otherRev = CurrentRevision(otherID)
    rows.Recordset.FindFirst "ReportID=" & id
    rows!Trait33 = "5.0": Check UXGridEdit("Trait33"), "Cell stage failed"
    Check CurrentRevision(id) = rev, "Cell committed report"
    Check DCount("*", "MeetingProposal", "SetID=" & f!txtMeeting & " AND ProposalState='Proposed'") = 1, "Proposal missing"
    rows!Trait33 = "4.0": Check UXGridEdit("Trait33"), "Second stage failed"
    Check DCount("*", "MeetingProposal", "SetID=" & f!txtMeeting & " AND ProposalState='Proposed'") = 1, "Duplicate field proposals"
    Check DCount("*", "MeetingProposal", "SetID=" & f!txtMeeting & " AND ProposalState='Superseded'") = 1, "Prior proposal not retained"
    Check InStr(f!uxGridAverage.Caption, "3.07") > 0, "Pending group average incorrect"
    rows!Trait34 = "9.0": Check Not UXGridEdit("Trait34"), "Invalid mark accepted"
    Check CurrentRevision(id) = rev, "Invalid mark saved report"
    rows.Recordset.FindFirst "ReportID=" & otherID
    rows!Trait34 = "4.0": Check UXGridEdit("Trait34"), "Second row stage failed"
    Check UXComparisonPage("Preview"), "Preview page failed"
    Check f!lstProposals.ListCount = 2 + IIf(f!lstProposals.ColumnHeads, 1, 0) And f!lstProposals.ColumnCount = 6, "Before/after preview missing"
    f!lstProposals.Selected(1) = True: f!lstProposals.Selected(2) = True
    ArchiveDraft otherID, True
    Check Not UXApplyComparison(), "Stale / archived batch accepted"
    Check CurrentRevision(id) = rev And CurrentRevision(otherID) = otherRev, "Failed batch partially saved"
    ArchiveDraft otherID, False
    Check UXApplyComparison(), "Comparison apply failed"
    Check CurrentRevision(id) <> rev And CurrentRevision(otherID) <> otherRev, "Atomic comparison apply failed"
    Check InStr(f!lblStatus.Caption, "applied atomically") > 0, "Comparison Results missing success message"
    Check Not f!cmdApply.Enabled, "Historical comparison allows another Apply"
    Check f!lstProposals.ListCount = 3 And f!lstProposals.Column(5, 1) = "Applied", "Applied before/after Results missing"
    Check DLookup("Trait33", "Report", "ReportID=" & id) = "4.0", "Wrong latest cell mark saved"
    rows.Recordset.FindFirst "ReportID=" & id
    Check Len(UXGridLockReason(id, rev)) > 0, "Stale row editable"
    cache = UXGridCacheFile()
    Check f!cmdBack.OnClick = "=UXCloseComparison()", "Return button does not close comparison"
    Check UXCloseComparison(), "Return to group failed"
    Check Not CurrentProject.AllForms("frmUXComparison").IsLoaded, "Comparison remained open on return"
    Check Len(Dir$(cache)) = 0, "Comparison cache retained after native close"
    Check DCount("*", "MSysObjects", "Name='PendingTraits'") = 0, "Workspace schema changed"
    For Each kind In Array("CHIEFEVAL", "FITREP")
        stage = CStr(kind) & " native cells and defaults"
        id = CommitDraft(0, CStr(kind), "Type test", "Type member", #1/1/2026#, #9/30/2026#, "", "Test", "")
        g = CreateSummaryGroup("Type " & kind, CStr(kind), IIf(kind = "FITREP", "O3", "E7"))
        AddGroupMember g, id, IIf(kind = "FITREP", "O3", "E7")
        If kind = "FITREP" Then
            otherID = CommitDraft(0, "FITREP", "NOB example", "NOB member", #1/1/2026#, #9/30/2026#, "", "Test", "")
            Set values = CreateObject("Scripting.Dictionary"): values("ObservationStatus") = "NOB"
            For Each key In Array("Trait33", "Trait34", "Trait35", "Trait36", "Trait37", "Trait38", "Trait39"): values(key) = "5.0": Next
            CommitSavedFields otherID, CurrentRevision(otherID), values, "NOB case"
            AddGroupMember g, otherID, "O3"
        End If
        Forms!frmMajorWorkbench!cboGroup = g
        Check UXOpenGroupDefaults(), "Typed defaults open failed"
        Check UXBulkSection(1), "Typed senior section failed"
        Check Forms!frmUXBulk!uxBlocks.Form!labelReportingSeniorName.Caption = UXFieldCaption("ReportingSeniorName", CStr(kind)), "Typed block caption wrong"
        UXCloseBlocks
        Check UXOpenComparison(), "Typed comparison failed"
        Set f = Forms!frmUXComparison: Set rows = f!uxTraitGrid.Form
        Check rows!heading33.Caption = UXFieldCaption("Trait33", CStr(kind)), "Trait headings wrong"
        rev = CurrentRevision(id)
        rows!Trait33 = "4.0": Check UXGridEdit("Trait33"), "Typed mark stage failed"
        Check CurrentRevision(id) = rev, "Typed edit saved without apply"
        If kind = "FITREP" Then Check InStr(f!uxGridAverage.Caption, "4.00") > 0 And InStr(f!uxGridAverage.Caption, "1 whole-report NOB excluded") > 0, "Whole-report NOB included in pending average"
        Check UXGridLoad(), "Proposal reload failed"
        Set rows = f!uxTraitGrid.Form
        Check rows!Trait33 = "4.0", "Staged value lost on reload"
        Set db = CurrentDb: Set query = db.CreateQueryDef("", f!lstProposals.RowSource)
        Check query.Parameters.Count = 0, "Preview has a parameter prompt": query.Close
        Check UXComparisonPage("Preview"), "Typed preview unavailable"
        Check Len(Nz(f!lstProposals.Column(3, 1), "")) = 0, "Before value wrong"
        Check f!lstProposals.Column(4, 1) = "4.0", "Full after value wrong"
        ArchiveDraft id, True: Check Len(UXGridLockReason(id, rev)) > 0, "Archive cell lock missing"
        rows!Trait34 = "4.0": Check Not UXGridEdit("Trait34"), "Archived row proposal accepted"
        ArchiveDraft id, False
        CurrentDb.Execute "UPDATE Report SET WorkflowState='Approved' WHERE ReportID=" & id
        Check Len(UXGridLockReason(id, rev)) > 0, "Approval cell lock missing"
        rows!Trait34 = "4.0": Check Not UXGridEdit("Trait34"), "Approved row proposal accepted"
        CurrentDb.Execute "UPDATE Report SET WorkflowState='Draft' WHERE ReportID=" & id
        DoCmd.Close acForm, "frmUXComparison", acSaveNo
    Next
    TestUXBlockTools = "PASS: " & checks & " shared blocks/defaults and editable comparison assertions": Exit Function
Failed:
    TestUXBlockTools = "FAIL: " & stage & ": " & Err.Number & " " & Err.Description
End Function
