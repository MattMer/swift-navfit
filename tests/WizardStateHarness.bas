Attribute VB_Name = "WizardStateHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal condition As Boolean, ByVal message As String)
 If Not condition Then Err.Raise vbObjectError + 9250, , message
 checks = checks + 1
End Sub
Private Function ApplyRejected() As Boolean
 On Error GoTo Rejected
 Call NAVFITApplyReturned
 Exit Function
Rejected:
 ApplyRejected = True
End Function
Private Function RejectLiveCancel(ByVal kind As String) As Boolean
 On Error GoTo Rejected
 If kind = "PDFImport" Then PDFClearPreview Else NAVFITClearReturned
 Exit Function
Rejected:
 RejectLiveCancel = (Err.Number = vbObjectError + 9203)
End Function
Private Function RejectContributorChange(ByVal id As Long) As Boolean
 On Error GoTo Rejected
 SwitchContributor id
 Exit Function
Rejected:
 RejectContributorChange = True
End Function
Private Function RejectOtherNote(ByVal identity As String, ByVal report As Long, ByVal version As Long) As Boolean
 On Error GoTo Rejected
 Call SaveReportMessageDraft(identity, report, version, "OTHER CONTRIBUTOR", 0)
 Exit Function
Rejected:
 RejectOtherNote = True
End Function
Public Function TestWizardTaskOrders(ByVal folder As String, ByVal template As String, ByVal pdfPath As String, Optional ByVal workspaceSource As String = "") As String
 Dim first As Long, second As Long, id As Long, note As String, otherNote As String, noteVersion As Long, revision As Long
 Dim aPDF As Object, aReturn As Object, bPDF As Object, bReturn As Object, rows As Collection, source As Object, choices As Object, token As String, result As String, stage As String
 On Error GoTo Failed
 checks = 0: stage = "prepare"
 result = PrepareWizardState(folder, template, pdfPath): Check Left$(result, 5) = "PASS:", result
 first = MajorAuthor(): id = CLng(DMin("ReportID", "Report")): revision = CurrentRevision(id)
 note = NewIdentity(): noteVersion = SaveReportMessageDraft(note, id, 0, "FIRST CONTRIBUTOR PRIVATE", 0)
 Set aPDF = WizardPreviewLoad("PDFImport"): Set aReturn = WizardPreviewLoad("NAVFITReturn")
 Check PDFClearPreview(False) And NAVFITClearReturned(False), "Live release failed"
 second = AddContributor("Second task-order contributor"): Check SwitchContributor(second), "Second contributor failed"
 Check RejectOtherNote(note, id, noteVersion), "Another contributor rewrote a private note draft"
 otherNote = NewIdentity(): Call SaveReportMessageDraft(otherNote, id, 0, "SECOND CONTRIBUTOR PRIVATE", 0)
 Set source = PreviewStructuredPDF(pdfPath): PDFCheckpointChoices "20", "Retain only", 0, "Fill empty fields"
 Set bPDF = WizardPreviewLoad("PDFImport")
 Set rows = NAVFITPreviewReturned(aReturn("SourcePath"), aReturn("BaselinePath"), 2000)
 NAVFITResolveReturned "OfficialComments", "Keep current": Set bReturn = WizardPreviewLoad("NAVFITReturn")
 Check PDFClearPreview(False) And NAVFITClearReturned(False), "Second live release failed"
 Check SwitchContributor(first), "First contributor failed"
 DoCmd.OpenForm "frmPDFImport": Check UXOpenNAVFITExchange(), "Return wizard did not open"
 Set source = PDFResumePreview(): Set rows = NAVFITResumeReturned()
 Set aPDF = WizardPreviewLoad("PDFImport"): Set aReturn = WizardPreviewLoad("NAVFITReturn")
 Check SwitchContributor(second), "Switch with stale first-contributor contexts failed"
 stage = "stale contributor PDF cancel"
 Check RejectLiveCancel("PDFImport"), "Stale contributor PDF cancel was accepted"
 Check Not UXCancelPDFImport(), "PDF cancel callback did not reject stale ownership"
 Check InStr(Forms!frmPDFImport!uxState.Caption, "Saved choices are retained") > 0, "PDF cancel callback omitted recovery guidance"
 Check WizardPreviewLoad("PDFImport")("StageGuid") = bPDF("StageGuid"), "Stale PDF cancel deleted the current contributor's saved choices"
 stage = "stale contributor return cancel"
 Check RejectLiveCancel("NAVFITReturn"), "Stale contributor return cancel was accepted"
 Check Not UXCancelNAVFITReturn(), "Return cancel callback did not reject stale ownership"
 Check InStr(Forms!frmNAVFITExchange!txtStatus, "Saved choices are retained") > 0, "Return cancel callback omitted recovery guidance"
 Check WizardPreviewLoad("NAVFITReturn")("StageGuid") = bReturn("StageGuid"), "Stale return cancel deleted the current contributor's saved choices"
 Set source = PDFResumePreview(): Check PDFClearPreview(), "Deliberate second PDF cancel failed"
 Set rows = NAVFITResumeReturned(): Check NAVFITClearReturned(), "Deliberate second return cancel failed"
 Check SwitchContributor(first), "Return to first contributor failed"
 Check WizardPreviewLoad("PDFImport")("StageGuid") = aPDF("StageGuid"), "Second contributor deleted first PDF choices"
 Check WizardPreviewLoad("NAVFITReturn")("StageGuid") = aReturn("StageGuid"), "Second contributor deleted first return choices"
 stage = "superseded same-contributor cancellation"
 Set source = PDFResumePreview(): Set aPDF = WizardPreviewLoad("PDFImport")
 token = WizardPreviewSave("PDFImport", aPDF("SourcePath"), aPDF("SourceHash"), "", "", 0, 0, aPDF("Choices"))
 Check RejectLiveCancel("PDFImport"), "Superseded PDF context cancelled a newer descriptor"
 Check WizardPreviewLoad("PDFImport")("StageGuid") = token, "Newer PDF descriptor was lost"
 Set source = PDFResumePreview(): Check PDFClearPreview(), "Explicit resumed PDF cancel failed"
 Set rows = NAVFITResumeReturned(): Set aReturn = WizardPreviewLoad("NAVFITReturn")
 token = WizardPreviewSave("NAVFITReturn", aReturn("SourcePath"), aReturn("SourceHash"), aReturn("BaselinePath"), aReturn("BaselineHash"), id, revision, aReturn("Choices"))
 Check RejectLiveCancel("NAVFITReturn"), "Superseded return context cancelled a newer descriptor"
 Check WizardPreviewLoad("NAVFITReturn")("StageGuid") = token, "Newer return descriptor was lost"
 Set rows = NAVFITResumeReturned()
 stage = "archive between preview and Apply"
 ArchiveDraft id, True
 Check ApplyRejected(), "Return Apply wrote an archived report"
 Check WizardPreviewLoad("NAVFITReturn")("StageGuid") <> "", "Rejected archived Apply lost saved choices"
 ArchiveDraft id, False: Check NAVFITClearReturned(), "Explicit return cancel failed"
 stage = "pending editor and contributor switch"
 DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
 Forms!frmEditor!txtStart = "raw invalid date": Check CheckpointEditor(), "Pending checkpoint failed"
 Check RejectContributorChange(second), "Contributor changed while editor remained open"
 Check MajorAuthor() = first And Forms!frmEditor!txtStart = "raw invalid date", "Rejected switch changed attribution or raw data"
 Check CurrentRevision(id) = revision And EditorHasPending(), "Rejected tasks committed or discarded pending values"
 Check UXLeaveEditor("Discard"), "Explicit task-order discard failed"
 Check DCount("*", "ReportNoteDraft") = 2 And DCount("*", "ReportMessage") = 0, "Task order published or lost private note drafts"
 Check DCount("*", "WizardPreview") = 0, "Explicit cancellations retained wizard rows"
 DoCmd.Close acForm, "frmPDFImport", acSaveNo: DoCmd.Close acForm, "frmNAVFITExchange", acSaveNo
 If workspaceSource <> "" Then
  stage = "workspace contributor cancellation"
  Check UXPreviewSwiftWorkspace(workspaceSource), "First workspace preview failed"
  Check SwitchContributor(second), "Second workspace contributor failed"
  Check UXPreviewSwiftWorkspace(workspaceSource), "Second workspace preview failed"
  Check SwitchContributor(first) And UXResumeSwiftWorkspace(), "First workspace resume failed"
  Check SwitchContributor(second), "Switch with stale workspace context failed"
  Set bReturn = WizardPreviewLoad("WorkspaceImport")
  Check Not UXCancelSwiftWorkspacePreview(), "Stale workspace cancel deleted another contributor's preview"
  Check WizardPreviewLoad("WorkspaceImport")("StageGuid") = bReturn("StageGuid"), "Stale workspace cancel changed saved choices"
  Check UXResumeSwiftWorkspace() And UXCancelSwiftWorkspacePreview(), "Explicit second workspace cancel failed"
  Check SwitchContributor(first) And UXResumeSwiftWorkspace(), "First workspace choices were lost"
  Set aReturn = WizardPreviewLoad("WorkspaceImport")
  token = WizardPreviewSave("WorkspaceImport", aReturn("SourcePath"), aReturn("SourceHash"), "", "", 0, 0, aReturn("Choices"))
  Check Not UXCancelSwiftWorkspacePreview(), "Superseded workspace context cancelled newer choices"
  Check WizardPreviewLoad("WorkspaceImport")("StageGuid") = token, "Newer workspace preview was lost"
  Check UXResumeSwiftWorkspace(), "Newer workspace resume failed"
  DoCmd.Close acForm, "frmUXWorkspaceImport", acSaveNo
  DoCmd.OpenForm "frmUXWorkspaceImport"
  Check Not UXApplySwiftWorkspace(), "Reopened workspace preview applied without Resume"
  Check InStr(Forms!frmUXWorkspaceImport!lblStatus.Caption, "ownership was reset") > 0, "Reopen failed for the wrong reason"
  Check UXResumeSwiftWorkspace() And UXCancelSwiftWorkspacePreview(), "Reopened workspace explicit cancellation failed"
  Check DCount("*", "WizardPreview") = 0, "Workspace cancellations left saved rows"
 End If
 TestWizardTaskOrders = "PASS: " & checks & " contributor/token cancellation, private drafts, archive/Apply and pending recovery assertions": Exit Function
Failed:
 TestWizardTaskOrders = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
Public Function PrepareWizardState(ByVal folder As String, ByVal template As String, Optional ByVal pdfPath As String = "") As String
 Dim id As Long, rows As Collection, db As DAO.Database, changes As Object, saved As Object, key As Variant, row As Variant, parts As Variant
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC WIZARD TEST", "Wizard tester"
 Set changes = CreateObject("Scripting.Dictionary")
 For Each row In Array(AdministrationKeys(), TraitKeys(), AssessmentKeys(), ReportContextKeys(), MemberDetailsKeys(), DecisionDetailsKeys())
  For Each key In row: changes(CStr(key)) = Null: Next
 Next
 For Each row In NAVFITMapRows("EVAL")
  parts = Split(CStr(row), "|")
  Select Case CStr(parts(2))
   Case "flag": changes(CStr(parts(0))) = "No"
   Case "trait": changes(CStr(parts(0))) = "3.0"
  End Select
 Next
 changes("ComponentActive") = "Yes": changes("ReportRegular") = "Yes": changes("OccasionPeriodic") = "Yes"
 changes("ObservationStatus") = "Observed": changes("PromotionRecommendation") = "Promotable"
 id = CommitDraft(0, "EVAL", "SYNTHETIC RETURN", "SYNTHETIC MEMBER", #1/1/2026#, #9/30/2026#, "Private synthetic notes", "Wizard seed", "SYNTHETIC ORIGINAL", administration:=FieldsSubset(changes, AdministrationKeys()), traits:=FieldsSubset(changes, TraitKeys()), assessment:=FieldsSubset(changes, AssessmentKeys()), reportContext:=FieldsSubset(changes, ReportContextKeys()), memberDetails:=FieldsSubset(changes, MemberDetailsKeys()), decisionDetails:=FieldsSubset(changes, DecisionDetailsKeys()))
 Call NAVFITExportReport(id, folder & "\export.accdb", template)
 FileCopy folder & "\export.accdb", folder & "\returned.accdb"
 Set db = DBEngine.OpenDatabase(folder & "\returned.accdb")
 db.Execute "UPDATE Reports SET Comments='SYNTHETIC RETURNED' WHERE ReportType Is Not Null", dbFailOnError: db.Close
 Set rows = NAVFITPreviewReturned(folder & "\returned.accdb", folder & "\export.accdb.swift-baseline.accdb", 2000)
 NAVFITResolveReturned "OfficialComments", "Keep current"
 Set saved = WizardPreviewLoad("NAVFITReturn")
 Check saved("Choices")("OfficialComments") = "Keep current", "Reviewed decision not staged"
 Check saved("SourceHash") = SourceFileHash(folder & "\returned.accdb"), "Source fingerprint not staged"
 Check saved("ExpectedRevisionID") = CurrentRevision(id), "Target revision not staged"
 Check DCount("*", "ReportRevision") = 1, "Preview changed a report"
 Dim pdf As Object
 If pdfPath = "" Then pdfPath = CurrentProject.Path & "\..\..\..\FORMS_REFERENCES_TEMPLATES\E1-E6 EVAL BLANK NAVPERS 1616-26 05-2025_Final.pdf"
 Set pdf = PreviewStructuredPDF(pdfPath)
 PDFCheckpointChoices "20", "Retain only", 0, "Fill empty fields"
 Check WizardPreviewLoad("PDFImport")("Choices")("Identifier") = "Retain only", "PDF mapping choice not staged"
 PrepareWizardState = "PASS: " & checks & " durable return preparation checks": Exit Function
Failed:
 PrepareWizardState = "FAIL: " & Err.Number & ": " & Err.Description
End Function
Public Function VerifyWizardReset(Optional ByVal includePDFUI As Boolean = False) As String
 Dim rows As Collection, saved As Object, original As String, source As DAO.Database, choices As Object, token As String, replacement As String, other As Long, rejected As Boolean
 On Error GoTo Failed
 Check ApplyRejected(), "Apply survived VBA reset without deliberate resume"
 Check DCount("*", "WizardPreview", "WizardKind='NAVFITReturn'") = 1, "VBA reset lost saved choices"
 Set saved = PDFResumePreview()
 Check saved("Choices")("Identifier") = "Retain only" And PDFImportKind() = "EVAL", "PDF reset resume lost source or mapping choice"
 If includePDFUI Then
  DoCmd.OpenForm "frmPDFImport"
  Check UXResumePDFImport(), "PDF Resume callback failed: " & Forms!frmPDFImport!uxState.Caption
  Check Forms!frmPDFImport!cboIdentifier = "Retain only" And Forms!frmPDFImport!cmdImport.Enabled, "PDF Resume did not restore reviewed mapping"
  Check DCount("*", "Report") = 1, "PDF Resume applied automatically"
  DoCmd.Close acForm, "frmPDFImport", acSaveNo
  Check DCount("*", "WizardPreview", "WizardKind='PDFImport'") = 1, "Closing PDF wizard lost saved choices"
 End If
 Check PDFClearPreview(), "PDF saved preview cancellation failed"
 Set saved = WizardPreviewLoad("NAVFITReturn")
 Check UXOpenNAVFITExchange(), "Return screen did not open after reset"
 Check UXResumeNAVFITReturn(), "Return Resume callback failed: " & Nz(Forms!frmNAVFITExchange!txtStatus, "")
 Check Forms!frmNAVFITExchange!lstChoices.RowSourceType = "Value List" And Forms!frmNAVFITExchange!lstChoices.ListCount > 75, "Resumed choice list did not populate"
 Check Forms!frmNAVFITExchange!cmdApply.Enabled, "Reviewed resumed return cannot be applied"
 DoCmd.Close acForm, "frmNAVFITExchange", acSaveNo
 Check ApplyRejected() And DCount("*", "WizardPreview", "WizardKind='NAVFITReturn'") = 1, "Closing return did not release live ownership while retaining saved choices"
 Set rows = NAVFITResumeReturned()
 Check NAVFITReturnedDecision("OfficialComments") = "Keep current", "Reset resume lost reviewed decision"
 Check DCount("*", "ReportRevision") = 1, "Resume applied automatically"
 other = AddContributor("Second wizard tester")
 Check SwitchContributor(other), "Contributor switch failed"
 Check ApplyRejected(), "Another contributor could apply an owned preview"
 Check SwitchContributor(CLng(DMin("UserProfileID", "UserProfile"))), "Original contributor switch failed"
 Set rows = NAVFITResumeReturned()
 Set source = DBEngine.OpenDatabase(saved("SourcePath"))
 source.Execute "UPDATE Reports SET Comments='SYNTHETIC CHANGED SOURCE' WHERE ReportType Is Not Null", dbFailOnError: source.Close
 Check ApplyRejected(), "Changed source could apply"
 Check DCount("*", "WizardPreview") = 1 And DCount("*", "ReportRevision") = 1, "Rejected source lost choices or wrote data"
 'Restoring semantic text does not restore the file hash: deliberately start a
 'new preview after this source mutation, retaining the prior choice explicitly.
 Set rows = NAVFITPreviewReturned(saved("SourcePath"), saved("BaselinePath"), 2000)
 NAVFITResolveReturned "OfficialComments", "Keep current"
 Set choices = CreateObject("Scripting.Dictionary"): choices.Add "Mode", "SYNTHETIC"
 token = WizardPreviewSave("PDFImport", saved("SourcePath"), SourceFileHash(saved("SourcePath")), "", "", 0, 0, choices)
 replacement = WizardPreviewSave("PDFImport", saved("SourcePath"), SourceFileHash(saved("SourcePath")), "", "", 0, 0, choices)
 On Error Resume Next
 WizardPreviewCheck "PDFImport", token: rejected = (Err.Number <> 0): Err.Clear
 On Error GoTo Failed
 Check rejected, "Superseded context token was accepted"
 Check WizardPreviewCancel("PDFImport"), "Explicit cancellation failed"
 Check DCount("*", "WizardPreview", "WizardKind='PDFImport'") = 0, "Cancellation retained owned preview"
 VerifyWizardReset = "PASS: " & checks & " reset, ownership, stale source and cancellation checks": Exit Function
Failed:
 VerifyWizardReset = "FAIL: " & Err.Number & ": " & Err.Description
End Function
Public Function TestPDFWizardApply(ByVal path As String) As String
 Dim saved As Object, id As Long, count As Long, hash As String
 On Error GoTo Failed
 checks = 0
 hash = SourceFileHash(path): count = DCount("*", "Report")
 Set saved = PreviewStructuredPDF(path)
 PDFCheckpointChoices "20", "Retain only", 0, "Fill empty fields"
 Check PDFClearPreview(False), "PDF live context release failed"
 Set saved = PDFResumePreview()
 Check saved("Choices")("Identifier") = "Retain only", "PDF apply lost reviewed identifier choice"
 id = ImportPDFPreview("20", "Retain only")
 Check id > 0 And DCount("*", "Report") = count + 1, "PDF resumed import did not create exactly one draft"
 Check DCount("*", "WizardPreview", "WizardKind='PDFImport'") = 0, "PDF apply retained an applicable preview"
 Check DCount("*", "WizardPreview", "WizardKind='NAVFITReturn'") = 1, "PDF apply deleted another wizard's preview"
 Check SourceFileHash(path) = hash, "PDF import changed source bytes"
 TestPDFWizardApply = "PASS: " & checks & " resumed PDF apply/isolation checks": Exit Function
Failed:
 TestPDFWizardApply = "FAIL: " & Err.Number & ": " & Err.Description
End Function
Public Function TestApplyOutcomeFault(ByVal folder As String, ByVal template As String, ByVal pdfPath As String, ByVal kind As String, ByVal fault As String) As String
 Dim result As String, reports As Long, revisions As Long, ok As Boolean, status As String, db As DAO.Database
 On Error GoTo Failed
 checks = 0
 result = PrepareWizardState(folder, template, pdfPath): Check Left$(result, 5) = "PASS:", result
 If kind = "PDFImport" Then
  DoCmd.OpenForm "frmPDFImport": Check UXResumePDFImport(), "PDF UI resume failed"
 Else
  Check UXOpenNAVFITExchange() And UXResumeNAVFITReturn(), "Return UI resume failed"
  NAVFITResolveReturned "OfficialComments", "Accept returned"
 End If
 reports = DCount("*", "Report"): revisions = DCount("*", "ReportRevision")
 Set db = CurrentDb
 If fault = "Rollback" Then db.TableDefs("ImportReceipt").Fields("AdapterVersion").ValidationRule = "='SYNTHETIC REJECT'"
 If kind = "PDFImport" Then ok = UXApplyPDFImport() Else ok = UXApplyNAVFITReturn()
 Check Not ok, "Fault was reported as success"
 If kind = "PDFImport" Then
  status = Forms!frmPDFImport!uxState.Caption
  Check Not Forms!frmPDFImport!cmdImport.Enabled, "Failed PDF Apply remained enabled"
 Else
  status = Forms!frmNAVFITExchange!txtStatus
  Check Not Forms!frmNAVFITExchange!cmdApply.Enabled, "Failed return Apply remained enabled"
 End If
 If fault = "Rollback" Then
  db.TableDefs("ImportReceipt").Fields("AdapterVersion").ValidationRule = ""
  Check InStr(status, "rolled back") > 0, "Confirmed rollback was not disclosed"
  Check DCount("*", "Report") = reports And DCount("*", "ReportRevision") = revisions, "Rollback retained imported data"
  Check DCount("*", "WizardPreview", "WizardKind='" & kind & "'") = 1, "Rollback lost reviewed choices"
  If kind = "PDFImport" Then
   Check UXResumePDFImport() And UXApplyPDFImport(), "PDF rollback retry failed"
  Else
   Check UXResumeNAVFITReturn(), "Return rollback resume failed"
   NAVFITResolveReturned "OfficialComments", "Accept returned"
   Check UXApplyNAVFITReturn(), "Return rollback retry failed"
  End If
 ElseIf fault = "Uncertain" Then
  Check InStr(status, "could not be confirmed") > 0, "Uncertain outcome was reported as a confirmed rollback"
  Check InStr(status, "Close and reopen") > 0, "Uncertain outcome omitted recovery guidance"
  Check DCount("*", "Report") = reports And DCount("*", "ReportRevision") = revisions, "Simulated uncertain outcome wrote data"
  Check DCount("*", "WizardPreview", "WizardKind='" & kind & "'") = 1, "Uncertain outcome deleted saved choices"
  TestApplyOutcomeFault = "PASS: " & checks & " " & kind & " " & fault & " outcome/recovery assertions": Exit Function
 Else
  Check InStr(status, "data was committed") > 0, "Committed data was described as unapplied"
  Check InStr(status, "rolled back") = 0 And InStr(status, "Nothing applied") = 0, "False rollback feedback after commit"
  Check InStr(status, "Close and reopen") > 0, "Committed failure omitted recovery guidance"
  Check DCount("*", "WizardPreview", "WizardKind='" & kind & "'") = 0, "Committed import left an applicable preview"
  If kind = "PDFImport" Then ok = UXApplyPDFImport() Else ok = UXApplyNAVFITReturn()
  Check Not ok, "Repeated Apply succeeded after committed failure"
 End If
 Check DCount("*", "ReportRevision") = revisions + IIf(kind = "PDFImport", 2, 1), "Apply/retry duplicated or lost its normal revision sequence"
 Check DCount("*", "ImportReceipt") = 1, "Apply/retry did not retain exactly one import receipt"
 Check DCount("*", "Report") = reports + IIf(kind = "PDFImport", 1, 0), "Apply/retry duplicated a report"
 TestApplyOutcomeFault = "PASS: " & checks & " " & kind & " " & fault & " outcome/exactly-once assertions": Exit Function
Failed:
 TestApplyOutcomeFault = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function WizardUnhandledFailure() As Boolean
 Err.Raise vbObjectError + 9251, "WizardUnhandledFailure", "SYNTHETIC unhandled wizard failure"
End Function
Public Function SeedWizardMigrationSource(ByVal folder As String, ByVal template As String) As String
 Dim rows As Collection
 On Error GoTo Failed
 Call NAVFITExportReport(1, folder & "\migration-export.accdb", template)
 FileCopy folder & "\migration-export.accdb", folder & "\migration-return.accdb"
 Set rows = NAVFITPreviewReturned(folder & "\migration-return.accdb", folder & "\migration-export.accdb.swift-baseline.accdb", 2000)
 NAVFITResolveReturned "OfficialComments", "Keep current"
 SeedWizardMigrationSource = "PASS: " & WizardPreviewLoad("NAVFITReturn")("StageGuid"): Exit Function
Failed:
 SeedWizardMigrationSource = "FAIL: " & Err.Number & ": " & Err.Description
End Function
Private Function RejectedWorkspaceContext(ByVal path As String, ByVal hash As String) As Object
 On Error GoTo Rejected
 Call ImportSwiftWorkspaceData(path, hash)
 Exit Function
Rejected:
 Set RejectedWorkspaceContext = SwiftLastFailureContext()
End Function
Public Function TestWizardMigration(ByVal path As String, ByVal expectedStage As String) As String
 Dim db As DAO.Database, rows As Collection, saved As Object, hash As String, count As Long
 Dim expectedFailure As Object, actualFailure As Object, key As Variant, progressBinding As String, restoreProgress As Boolean
 On Error GoTo Failed
 Check SwitchContributor(CLng(DMin("UserProfileID", "UserProfile"))), "Bootstrap confirmation failed"
 hash = SourceFileHash(path)
 Check Not UXApplySwiftWorkspace(), "Workspace Apply survived reset without deliberate Resume"
 Check UXResumeSwiftWorkspace(), "Workspace preview did not survive reset"
 Set saved = WizardPreviewLoad("WorkspaceImport")
 Check saved("SourceHash") = hash, "Workspace source fingerprint missing"
 Set db = CurrentDb
 db.Execute "UPDATE Workspace SET CommandName='SYNTHETIC CHANGED BOOTSTRAP'", dbFailOnError
 Check Not UXApplySwiftWorkspace(), "Changed workspace setup could apply"
 Check DCount("*", "WizardPreview", "WizardKind='WorkspaceImport'") = 1, "Changed setup discarded saved preview"
 db.Execute "UPDATE Workspace SET CommandName='SYNTHETIC BOOTSTRAP'", dbFailOnError
 Check UXResumeSwiftWorkspace(), "Restored setup could not resume deliberately"
 db.TableDefs("WizardPreview").Fields("ChoiceXML").ValidationRule = "='FORCED LATE REJECTION'"
 'Establish the native primary independently of UI faults, on the same target.
 Set expectedFailure = RejectedWorkspaceContext(path, hash)
 Check Not expectedFailure Is Nothing, "Late rejection did not produce native diagnostics"
 Check expectedFailure("Transaction") = "Rolled back", "Direct late rejection did not confirm rollback"
 progressBinding = Forms!frmUXWorkspaceImport!txtProgress.ControlSource: restoreProgress = True
 'A calculated textbox rejects progress assignments without production test hooks.
 Forms!frmUXWorkspaceImport!txtProgress.ControlSource = "='Synthetic progress paint failure'"
 Check Not UXApplySwiftWorkspace(), "Late staged-row failure was accepted"
 Set actualFailure = SwiftLastFailureContext()
 For Each key In Array("Number", "Description", "Source", "Stage", "Transaction")
  Check actualFailure(key) = expectedFailure(key), "Progress/cleanup replaced primary " & key
 Next
 Check InStr(actualFailure("Cleanup"), "Paint progress (") > 0, "Optional progress failure was not recorded as secondary"
 Check InStr(Forms!frmUXWorkspaceImport!lblStatus.Caption, "Diagnostic: " & actualFailure("Identity")) > 0, "Primary import diagnostic missing from feedback"
 Check Not UXWorkspaceImportBusy() And Not Forms!frmUXWorkspaceImport.Modal, "Failed import left UI locked"
 Check Forms!frmUXWorkspaceImport!cmdApply.Enabled And Forms!frmUXWorkspaceImport!cmdClose.Enabled, "Confirmed rollback did not restore retry/close"
 Check DCount("*", "Report") = 0, "Late failure did not roll back copied reports"
 Check DLookup("CommandName", "Workspace") = "SYNTHETIC BOOTSTRAP", "Late failure replaced bootstrap"
 Check DCount("*", "WizardPreview", "WizardKind='WorkspaceImport'") = 1, "Late failure lost the import descriptor"
 db.TableDefs("WizardPreview").Fields("ChoiceXML").ValidationRule = "": db.Close
 Check UXResumeSwiftWorkspace(), "Workspace preview could not resume after rollback"
 'Keep painting broken during retry: optional display faults must not fail data import.
 Check UXApplySwiftWorkspace(), "Workspace staged import retry failed"
 Set actualFailure = SwiftLastFailureContext()
 Check actualFailure("Transaction") = "Committed", "Retry display diagnostic lost committed outcome"
 Check InStr(actualFailure("Cleanup"), "Paint progress (") > 0, "Retry did not exercise optional paint failure"
 Check Not UXWorkspaceImportBusy() And Not Forms!frmUXWorkspaceImport.Modal, "Retry left importer busy/modal"
 Check Not Forms!frmUXWorkspaceImport!cmdApply.Enabled And Forms!frmUXWorkspaceImport!cmdClose.Enabled, "Committed retry left wrong action state"
 Forms!frmUXWorkspaceImport!txtProgress.ControlSource = progressBinding: restoreProgress = False
 Check DCount("*", "Report") = 1 And DCount("*", "ReportRevision") = 3, "Migration changed report history"
 Check SourceFileHash(path) = hash, "Workspace import changed source bytes"
 Check SwitchContributor(CLng(DMin("UserProfileID", "UserProfile"))), "Imported contributor confirmation failed"
 Set saved = WizardPreviewLoad("NAVFITReturn")
 Check saved("StageGuid") = expectedStage, "Migration changed saved preview identity"
 Set rows = NAVFITResumeReturned()
 Check NAVFITReturnedDecision("OfficialComments") = "Keep current", "Imported preview lost reviewed choice"
 Check DCount("*", "WizardPreview", "WizardKind='WorkspaceImport'") = 0, "Committed import retained its target descriptor"
 TestWizardMigration = "PASS: " & checks & " staged-state migration, late rollback and retry checks": Exit Function
Failed:
 TestWizardMigration = "FAIL: " & Err.Number & ": " & Err.Description
 On Error Resume Next
 If restoreProgress Then Forms!frmUXWorkspaceImport!txtProgress.ControlSource = progressBinding
 On Error GoTo 0
End Function
Public Function PrepareWorkspaceResume(ByVal path As String) As Boolean
 ConfigureWorkspace "SYNTHETIC BOOTSTRAP", "Bootstrap tester"
 PrepareWorkspaceResume = UXPreviewSwiftWorkspace(path)
End Function
Public Function VerifyWizardRestart() As String
 Dim rows As Collection, id As Long, saved As Object, changes As Object, count As Long
 On Error GoTo Failed
 Check ApplyRejected(), "Restart applied an unresumed preview"
 Check SwitchContributor(CLng(DMin("UserProfileID", "UserProfile"))), "Restart contributor confirmation failed"
 Set saved = WizardPreviewLoad("NAVFITReturn")
 Set rows = NAVFITResumeReturned()
 Check NAVFITReturnedDecision("OfficialComments") = "Keep current", "Restart lost reviewed choice"
 Check DCount("*", "ReportRevision") = 1, "Restart resume applied values"
 id = saved("TargetReportID")
 Call CommitDraft(id, "EVAL", "SYNTHETIC CHANGED TARGET", "SYNTHETIC MEMBER", #1/1/2026#, #9/30/2026#, "Private synthetic notes", "Target stale test", "SYNTHETIC CURRENT")
 Check ApplyRejected(), "Changed target revision could apply"
 Check DCount("*", "WizardPreview") = 1, "Stale target lost saved choices"
 Set rows = NAVFITPreviewReturned(saved("SourcePath"), saved("BaselinePath"), 2000)
 NAVFITResolveReturned "OfficialComments", "Accept returned"
 count = DCount("*", "ReportRevision")
 Check NAVFITApplyReturned() = id, "Reviewed return apply failed"
 Check DCount("*", "ReportRevision") = count + 1, "Return did not create one revision"
 Check DCount("*", "WizardPreview") = 0, "Applied return left a resumable preview"
 Check ApplyRejected(), "Applied preview could be applied twice"
 VerifyWizardRestart = "PASS: " & checks & " restart, stale target and exactly-once apply checks": Exit Function
Failed:
 VerifyWizardRestart = "FAIL: " & Err.Number & ": " & Err.Description
End Function
