Attribute VB_Name = "UXDemoHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal value As Boolean, ByVal message As String)
    If Not value Then Err.Raise vbObjectError + 5300, , message
    checks = checks + 1
End Sub
Private Function RejectSeed() As Boolean
    On Error GoTo Rejected
    Call SeedUXDemo()
    Exit Function
Rejected:
    RejectSeed = (Err.Number = vbObjectError + 5100)
End Function
Public Function TestUXDemo() As String
    Dim kind As Variant, id As Long, revision As Long, f As Access.Form, s As Access.Form, groupID As Long, average As Object, stage As String
    On Error GoTo Failed
    stage = "dataset"
    Check DCount("*", "Report") = 9 And DCount("*", "SummaryGroup") = 3, "Synthetic roster incomplete"
    Check DCount("*", "UserProfile") = 2 And DCount("*", "ReportingSeniorRecord") = 1 And DCount("*", "RSCABaseline") = 3, "Synthetic directory incomplete"
    Check DCount("*", "Report", "IsArchived=True") = 1, "Archive example missing"
    Check DCount("*", "Report", "WorkflowState='Approved'") = 1, "Approval example missing"
    Check DCount("*", "Report", "ObservationStatus='NOB'") = 1, "NOB example missing"
    Check DCount("*", "RecoveryDraft") = 1, "Recovery example missing"
    Check RejectSeed(), "Demo seeder accepted a populated workspace"
    Check DCount("*", "Report") = 9, "Rejected seed changed roster"
    stage = "overview": DoCmd.OpenForm "frmHome"
    Check UXDashboardRefresh(), "Demo dashboard failed"
    Check CLng(Forms!frmHome!uxReportCount.Caption) = 8 And CLng(Forms!frmHome!uxGroupCount.Caption) = 3, "Demo metrics incorrect"
    Check CLng(Forms!frmHome!uxWorkCount.Caption) = 2 And Forms!frmHome!cmdRecovery.Enabled, "Work/recovery dashboard actions unavailable"
    stage = "report types"
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        id = DLookup("ReportID", "Report", "FormKind=" & QuoteSQL(CStr(kind)) & " AND IsArchived=False AND WorkflowState='Draft' AND ObservationStatus='Observed'")
        revision = CurrentRevision(id)
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Set f = Forms!frmEditor: Set s = f!uxFieldSurface.Form
        Check UXSelectSection(5), "Demo report section failed"
        Check s!labelOfficialComments.Caption Like IIf(kind = "EVAL", "43.*", IIf(kind = "CHIEFEVAL", "40.*", "41.*")), "Demo comments label incorrect"
        Check UXSaveEditor(), "Valid demo report could not save"
        Check CurrentRevision(id) = revision, "No-op demo save appended a revision"
        DoCmd.Close acForm, "frmEditor", acSaveNo
    Next
    stage = "locked report": id = DLookup("ReportID", "Report", "WorkflowState='Approved'")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Call UXSelectSection(5)
    Check Not Forms!frmEditor!cmdSave.Enabled And Forms!frmEditor!uxFieldSurface.Form!txtOfficialComments.Locked, "Approved demo report is editable"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    stage = "groups": DoCmd.OpenForm "frmUXGroups"
    groupID = DLookup("GroupID", "SummaryGroup", "Paygrade='E5'")
    Forms!frmUXGroups!lstGroups = groupID
    Check CLng(Forms!frmUXGroups!lstGroups.Column(4)) = DCount("*", "GroupMember", "GroupID=" & groupID), "Demo group report count incorrect"
    Check UXOpenGroup(), "Demo group failed to open"
    Check InStr(Forms!frmMajorWorkbench!uxGroupMeta.Caption, "PSG RSCA") > 0, "Group context averages missing"
    Set average = RSGroupPSG(groupID)
    Check Not IsNull(average("Value")) And Not average("Estimated"), "Exact demo PSG unavailable"
    Call UXWorkbenchMode("Comparison")
    Check UXOpenComparison(), "Demo comparison failed"
    Check Forms!frmUXComparison!lstTraits.ListCount >= 3, "Demo comparison roster empty"
    Check UXComparisonPage("Preview") And Forms!frmUXComparison!lstProposals.Visible, "Comparison preview unavailable"
    DoCmd.Close acForm, "frmUXComparison", acSaveNo
    stage = "my work": DoCmd.OpenForm "frmMyWork"
    Check Forms!frmMyWork!lstTasks.ListCount = 2, "Demo assigned work list incorrect"
    Check UXMyWorkMode("Contributors") And Forms!frmMyWork!cboContributor.ListCount = 2, "Demo contributor switching unavailable"
    TestUXDemo = "PASS: " & checks & " populated native demo assertions": Exit Function
Failed:
    TestUXDemo = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
