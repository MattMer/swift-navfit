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
Public Function TestWizardMigration(ByVal path As String, ByVal expectedStage As String) As String
 Dim db As DAO.Database, rows As Collection, saved As Object, hash As String, count As Long
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
 Check Not UXApplySwiftWorkspace(), "Late staged-row failure was accepted"
 Check DCount("*", "Report") = 0, "Late failure did not roll back copied reports"
 Check DLookup("CommandName", "Workspace") = "SYNTHETIC BOOTSTRAP", "Late failure replaced bootstrap"
 Check DCount("*", "WizardPreview", "WizardKind='WorkspaceImport'") = 1, "Late failure lost the import descriptor"
 db.TableDefs("WizardPreview").Fields("ChoiceXML").ValidationRule = "": db.Close
 Check UXResumeSwiftWorkspace(), "Workspace preview could not resume after rollback"
 Check UXApplySwiftWorkspace(), "Workspace staged import retry failed"
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
