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
    Dim kind As Variant, id As Long, revision As Long, f As Access.Form, s As Access.Form, groupID As Long, average As Object, stage As String, rows As DAO.Recordset
    On Error GoTo Failed
    stage = "dataset"
    Check DCount("*", "Report") = 13 And DCount("*", "SummaryGroup") = 5, "Synthetic roster incomplete"
    Check DCount("*", "UserProfile") = 2 And DCount("*", "ReportingSeniorRecord") = 1 And DCount("*", "RSCABaseline") = 4, "Synthetic directory incomplete"
    Check DCount("*", "Report", "IsArchived=True") = 1, "Archive example missing"
    Check DCount("*", "Report", "WorkflowState='Approved'") = 1, "Approval example missing"
    Check DCount("*", "Report", "ObservationStatus='NOB'") = 1, "NOB example missing"
    Check DCount("*", "RecoveryDraft") = 1, "Recovery example missing"
    Check RejectSeed(), "Demo seeder accepted a populated workspace"
    Check DCount("*", "Report") = 13, "Rejected seed changed roster"
    Check DemoReportIssues("DEMO - Clean reports (0 report issues)") = 0, "Clean group member issues"
    Check DemoReportIssues("DEMO - Practice corrections (3 report issues)") = 3, "Practice group member issue count"
    groupID = DLookup("GroupID", "SummaryGroup", "Title='DEMO - Clean reports (0 report issues)'")
    Set average = RSGroupPSG(groupID)
    Check Not IsNull(average("Value")), "Clean E6 group lacks matching RSCA baseline"
    Check Abs(average("Value") - 476# / 114) < 0.00001, "Clean E6 PSG arithmetic incorrect"
    stage = "overview": DoCmd.OpenForm "frmHome"
    Check UXDashboardRefresh(), "Demo dashboard failed"
    Check CLng(Forms!frmHome!uxReportCount.Caption) = 12 And CLng(Forms!frmHome!uxGroupCount.Caption) = 5, "Demo metrics incorrect"
    Check CLng(Forms!frmHome!uxWorkCount.Caption) = UXWorkCount() And Forms!frmHome!cmdRecovery.Enabled, "Work/recovery dashboard actions unavailable"
    stage = "report types"
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        id = DLookup("ReportID", "Report", "FormKind=" & QuoteSQL(CStr(kind)) & " AND IsArchived=False AND WorkflowState='Draft' AND ObservationStatus='Observed'")
        revision = CurrentRevision(id)
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Set f = Forms!frmEditor: Set s = f!uxFieldSurface.Form
        Check f.OnUnload = "[Event Procedure]", "Demo seeding removed native close protection"
        Check UXSelectSection(5), "Demo report section failed"
        Check s!labelOfficialComments.Caption Like IIf(kind = "EVAL", "43.*", IIf(kind = "CHIEFEVAL", "40.*", "41.*")), "Demo comments label incorrect"
        Check UXSaveEditor(), "Valid demo report could not save"
        Check CurrentRevision(id) = revision, "No-op demo save appended a revision"
        Call UXLeaveEditor("Discard")
    Next
    stage = "locked report": id = DLookup("ReportID", "Report", "WorkflowState='Approved'")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Call UXSelectSection(5)
    Check Not Forms!frmEditor!cmdSave.Enabled And Forms!frmEditor!uxFieldSurface.Form!txtOfficialComments.Locked, "Approved demo report is editable"
    Call UXLeaveEditor("Discard")
    stage = "groups": DoCmd.OpenForm "frmUXGroups"
    groupID = DLookup("GroupID", "SummaryGroup", "Paygrade='E5'")
    Forms!frmUXGroups!lstGroups = groupID
    Check CLng(Forms!frmUXGroups!lstGroups.Column(4)) = DCount("*", "GroupMember", "GroupID=" & groupID), "Demo group report count incorrect"
    Check UXOpenGroup(), "Demo group failed to open"
    Check Forms!frmUXComparison!uxGroupContext = groupID And Forms!frmUXComparison!uxSeniorCaption.Visible, "Live group context or averages missing"
    Set average = RSGroupPSG(groupID)
    Check Not IsNull(average("Value")) And Not average("Estimated"), "Exact demo PSG unavailable"
    Set rows = Forms!frmUXComparison!uxTraitGrid.Form.RecordsetClone
    rows.MoveLast: Check rows.RecordCount >= 3, "Demo live roster empty": rows.Close
    Check UXComparisonPage("Preview") And Forms!frmUXComparison!lstProposals.Visible, "Comparison preview unavailable"
    DoCmd.Close acForm, "frmUXComparison", acSaveNo
    stage = "my work": DoCmd.OpenForm "frmMyWork"
    Check Forms!frmMyWork!lstTasks.ListCount = 2, "Demo assigned work list incorrect"
    Check UXMyWorkMode("Contributors") And Forms!frmMyWork!lstContributors.ListCount = 2, "Demo contributor switching unavailable"
    TestUXDemo = "PASS: " & checks & " populated native demo assertions": Exit Function
Failed:
    TestUXDemo = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function

Private Function DemoReportIssues(ByVal title As String) As Long
 Dim groupID As Long, r As DAO.Recordset, findings As Collection
 groupID = DLookup("GroupID", "SummaryGroup", "Title=" & QuoteSQL(title))
 Set r = CurrentDb.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & groupID, dbOpenSnapshot)
 Do While Not r.EOF
  Set findings = ValidateLiveFields(SavedFields(r!ReportID), r!ReportID)
  DemoReportIssues = DemoReportIssues + findings.Count
  r.MoveNext
 Loop
 r.Close
End Function
Public Function DemoValidationSummary() As String
 Dim title As Variant, groupID As Long, findings As Collection, text As String
 For Each title In Array("DEMO - Clean reports (0 report issues)", "DEMO - Practice corrections (3 report issues)")
  groupID = DLookup("GroupID", "SummaryGroup", "Title=" & QuoteSQL(CStr(title)))
  Set findings = ValidateLiveGroup(groupID)
  text = text & title & ": " & DemoReportIssues(CStr(title)) & " member issues; " & findings.Count & " total findings" & vbCrLf
 Next
 DemoValidationSummary = text
End Function