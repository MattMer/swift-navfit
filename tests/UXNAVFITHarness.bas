Attribute VB_Name = "UXNAVFITHarness"
Option Compare Binary
Option Explicit
Private checks As Long
Private Sub Check(ByVal result As Boolean, ByVal message As String)
    checks = checks + 1
    If Not result Then Err.Raise vbObjectError + 5600, , message
End Sub
Public Function TestUXNAVFIT(ByVal source As String, ByVal template As String) As String
    Dim before As Long, id As Long, file As String, changes As Object, db As DAO.Database, groupID As Long, phase As String
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC NAVFIT UI", "Synthetic tester"
    Call OpenReviewFile(source)
    Check Forms!frmImportReview!cmdPreview.OnClick = "=UXPreviewFullNAVFIT()", "Full preview not wired"
    Check Forms!frmImportReview!cmdImport.OnClick = "=UXApplyFullNAVFIT()", "Full apply not wired"
    Check Not Forms!frmImportReview!cmdImport.Enabled, "Apply enabled before preview"
    Forms!frmImportReview!lstImportReports = 2: Call ReviewReportFields
    Check UXPreviewFullNAVFIT(), "Automatic 20xx preview failed"
    Forms!frmImportReview!txtImportMember = "Synthetic local display"
    before = DCount("*", "Report")
    Check UXPreviewFullNAVFIT(), "Full preview failed"
    Check DCount("*", "Report") = before, "Preview wrote a report"
    Check InStr(Forms!frmImportReview!uxPreview, "128") > 0 And InStr(Forms!frmImportReview!uxPreview, "After:") > 0, "Full mapping scope missing"
    Check InStr(Forms!frmImportReview!uxPreview, "Synthetic local display") > 0, "Display-name preview absent"
    Forms!frmImportReview!txtImportMember = "Changed unpreviewed display"
    Check Not UXApplyFullNAVFIT() And DCount("*", "Report") = before, "Changed display accepted"
    Check UXPreviewFullNAVFIT(), "Fresh full preview failed"
    Check UXApplyFullNAVFIT(), "Full apply failed"
    id = DMax("ReportID", "Report")
    Check SavedFields(id)("MemberName") = "Changed unpreviewed display", "Reviewed display lost"
    Check SavedFields(id)("OfficialName") <> SavedFields(id)("MemberName"), "Local display overwrote official name"
    Check Not UXApplyFullNAVFIT() And DCount("*", "Report") = before + 1, "Repeated apply wrote report"
    Check UXOpenFullNAVFITResult(), "Result cannot open"
    Check CLng(Forms!frmEditor!txtReportID) = id, "Wrong report opened"
    Call UXLeaveEditor
    Call CloseReviewFile
    Check UXOpenNAVFITExchange(), "Exchange workspace cannot open"
    Check Not Forms!frmNAVFITExchange!cmdApply.Enabled, "Return Apply enabled without preview"
    Check Forms!frmNAVFITExchange!txtReturned.AfterUpdate = "=UXNAVFITInputChanged()", "Changed returned path retains stale preview"
    Check Forms!frmNAVFITExchange.OnUnload = "=NAVFITClearReturned()", "Preview state survives closing"
    file = source & ".ui-export.accdb"
    Call NAVFITExportReport(id, file, template)
    Set db = DBEngine.OpenDatabase(file)
    db.Execute "UPDATE Reports SET Comments='Native return UI conflict' WHERE ReportType Is Not Null", dbFailOnError: db.Close
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "OfficialComments", "Keep this local UI value"
    Call CommitSavedFields(id, CurrentRevision(id), changes, "UI conflict test")
    Forms!frmNAVFITExchange!txtReturned = file
    Forms!frmNAVFITExchange!txtBaseline = file & ".swift-baseline.accdb"
    Check UXPreviewNAVFITReturn(), "Return UI preview failed"
    Forms!frmNAVFITExchange!lstChoices = "OfficialComments"
    Check UXNAVFITReturnDetail(), "Return detail unavailable"
    Check Forms!frmNAVFITExchange!lstChoices.Column(2) = "Conflict", "Conflict row missing"
    Check UXResolveNAVFITReturn("Keep current"), "UI resolution failed"
    Check Forms!frmNAVFITExchange!lstChoices.Column(2) = "Keep current", "Resolved row still shows Conflict"
    Check InStr(Forms!frmNAVFITExchange!txtDetail, "Keep this local UI value") > 0, "Full local value absent"
    Check UXApplyNAVFITReturn(), "Resolved UI return cannot apply"
    Check SavedFields(id)("OfficialComments") = "Keep this local UI value", "UI return lost local value"
    phase = "Create group": groupID = CreateSummaryGroup("Synthetic UI export group", "FITREP", "O3")
    phase = "Add member"
    AddGroupMember groupID, id, "O3"
    phase = "Set export controls"
    Forms!frmNAVFITExchange!cboGroup = groupID
    Forms!frmNAVFITExchange!txtTemplate = template
    Forms!frmNAVFITExchange!txtExportPath = source & ".ui-group.accdb"
    Check UXPreviewNAVFITGroup(), "Group UI preview failed"
    Check Forms!frmNAVFITExchange!cmdExportGroup.Enabled, "Reviewed group export unavailable"
    Check InStr(Forms!frmNAVFITExchange!txtStatus, "Saved members:") > 0, "Group roster preview missing"
    Check UXNAVFITInputChanged() And Not Forms!frmNAVFITExchange!cmdExportGroup.Enabled, "Group preview not invalidated"
    Check Not UXExportNAVFITGroup(), "Unpreviewed group exported"
    Check UXPreviewNAVFITGroup() And UXExportNAVFITGroup(), "Reviewed group cannot export"
    Check Len(Dir$(Forms!frmNAVFITExchange!txtExportPath & ".swift-baseline.accdb")) > 0, "Group baseline absent at the actual generated export path"
    TestUXNAVFIT = "PASS: " & checks & " NAVFIT UI assertions": Exit Function
Failed:
    TestUXNAVFIT = "FAIL after " & checks & " checks (" & phase & "): " & Err.Number & " " & Err.Source & " " & Err.Description
End Function
