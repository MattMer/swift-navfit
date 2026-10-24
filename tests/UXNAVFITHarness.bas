Attribute VB_Name = "UXNAVFITHarness"
Option Compare Binary
Option Explicit
Private checks As Long
Private Sub Check(ByVal result As Boolean, ByVal message As String)
    checks = checks + 1
    If Not result Then Err.Raise vbObjectError + 5600, , message
End Sub
Public Function TestUXNAVFIT(ByVal source As String, ByVal template As String) As String
    Dim before As Long, id As Long, file As String, changes As Object, db As DAO.Database, groupID As Long, phase As String, revision As Long, recovery As Long
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
    revision = CurrentRevision(id): Forms!frmEditor!txtNotes = "SYNTHETIC pending export test"
    Check CheckpointEditor(), "Pending checkpoint failed": recovery = DMax("RecoveryDraftID", "RecoveryDraft")
    Check NAVFITExportReport(id, source & ".pending-export.accdb", template) <> "", "Open editor blocked saved-only export"
    Check CurrentProject.AllForms("frmEditor").IsLoaded And EditorHasPending(), "Export closed/cleared pending editor"
    Check Forms!frmEditor!txtNotes = "SYNTHETIC pending export test" And CurrentRevision(id) = revision, "Export saved/discarded pending text"
    Check DLookup("DraftTitle", "RecoveryDraft", "RecoveryDraftID=" & recovery) <> "", "Export removed recovery"
    Check Not UXLeaveEditor("Cancel"), "Export bypassed explicit navigation choice"
    Check UXLeaveEditor("Discard"), "Explicit test cleanup failed"
    Call CloseReviewFile
    Check UXOpenNAVFITExchange(), "Exchange workspace cannot open"
    Check Not Forms!frmNAVFITExchange!cmdApply.Enabled, "Return Apply enabled without preview"
    Check Forms!frmNAVFITExchange!txtReturned.AfterUpdate = "=UXNAVFITInputChanged()", "Changed returned path retains stale preview"
    Check Forms!frmNAVFITExchange.OnUnload = "=NAVFITClearReturned(False)", "Preview state survives closing"
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
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "PromotionRecommendation", Null: changes.Add "PeriodEnd", Null: changes.Add "ReportingSeniorSSN", Null
    Call CommitSavedFields(id, CurrentRevision(id), changes, "Synthetic incomplete UI export")
    CurrentDb.Execute "UPDATE Report SET IsArchived=True,WorkflowState='Approved' WHERE ReportID=" & id, dbFailOnError
    CurrentDb.Execute "UPDATE SummaryGroup SET IsArchived=True WHERE GroupID=" & groupID, dbFailOnError
    Check UXOpenNAVFITExchange(), "Archived export screen failed"
    Check InStr(Forms!frmNAVFITExchange!cboReport.RowSource, "WHERE IsArchived=False") = 0 And InStr(Forms!frmNAVFITExchange!cboGroup.RowSource, "WHERE IsArchived=False") = 0, "Archived exports hidden"
    phase = "Set export controls"
    Forms!frmNAVFITExchange!cboGroup = groupID
    Forms!frmNAVFITExchange!txtTemplate = template
    Forms!frmNAVFITExchange!cboReport = id
    Forms!frmNAVFITExchange!txtExportPath = source & ".incomplete-ui-report.accdb"
    If Not UXExportNAVFIT() Then Err.Raise vbObjectError + 5600, "UXNAVFITHarness", "Issue-bearing UI report blocked: " & Forms!frmNAVFITExchange!txtStatus
    Check True, "Issue-bearing UI report exported"
    Check InStr(Forms!frmNAVFITExchange!txtStatus, "Findings do not block export") > 0, "Informational export notice missing"
    Check InStr(Forms!frmNAVFITExchange!txtExportPath, "UNSPECIFIED.accdb") > 0, "Missing end date blocked safe naming"
    Forms!frmNAVFITExchange!txtExportPath = source & ".ui-group.accdb"
    Forms!frmNAVFITExchange!chkPSG = True
    Check UXPreviewNAVFITGroup(), "Group UI preview failed"
    Check Forms!frmNAVFITExchange!cmdExportGroup.Enabled, "Reviewed group export unavailable"
    Check InStr(Forms!frmNAVFITExchange!txtStatus, "Saved members:") > 0, "Group roster preview missing"
    Check InStr(Forms!frmNAVFITExchange!txtStatus, "No promotion selected: 1") > 0 And InStr(Forms!frmNAVFITExchange!txtStatus, "PSG RSCA unavailable") > 0, "Incomplete preview/fallback missing"
    Check UXNAVFITInputChanged() And Not Forms!frmNAVFITExchange!cmdExportGroup.Enabled, "Group preview not invalidated"
    Check Not UXExportNAVFITGroup(), "Unpreviewed group exported"
    Check UXPreviewNAVFITGroup() And UXExportNAVFITGroup(), "Reviewed group cannot export"
    Check Len(Dir$(Forms!frmNAVFITExchange!txtExportPath & ".swift-baseline.accdb")) > 0, "Group baseline absent at the actual generated export path"
    TestUXNAVFIT = "PASS: " & checks & " NAVFIT UI assertions": Exit Function
Failed:
    TestUXNAVFIT = "FAIL after " & checks & " checks (" & phase & "): " & Err.Number & " " & Err.Source & " " & Err.Description
End Function
