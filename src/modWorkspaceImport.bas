Attribute VB_Name = "modWorkspaceImport"
Option Explicit
Private importPath As String, importHash As String
Private importing As Boolean, importStarted As Date, progressText As String, lastProgressPaint As Single
Private pulseBusy As Boolean, enabledBefore As Object, modalBefore As Boolean
Private displayFailure As Object
'Durable descriptors contain fingerprints and reviewed choices, never executable
'source content. Live extracted values are rebuilt only by deliberate Resume.
Private workspacePreview As Object
Public Function WizardPreviewOwner() As Long
 If DCount("*", "Workspace") > 0 Then WizardPreviewOwner = ActiveContributorID()
End Function
Private Function WizardWhere(ByVal kind As String) As String
 Dim owner As Long
 If kind <> "NAVFITReturn" And kind <> "PDFImport" And kind <> "WorkspaceImport" Then Err.Raise 5, "WizardWhere", "Unknown preview kind."
 owner = WizardPreviewOwner()
 WizardWhere = "WizardKind=" & QuoteSQL(kind) & " AND "
 WizardWhere = WizardWhere & "Nz(OwnerProfileID,0)=" & owner
End Function
Public Function WizardPreviewLoad(ByVal kind As String) As Object
 Dim r As DAO.Recordset, field As DAO.Field, result As Object, dom As Object, node As Object, choices As Object
 Set r = CurrentDb.OpenRecordset("SELECT * FROM WizardPreview WHERE " & WizardWhere(kind), dbOpenSnapshot)
 If r.EOF Then r.Close: Err.Raise vbObjectError + 9200, "WizardPreviewLoad", "No saved preview for this contributor. Choose and preview a source first."
 Set result = CreateObject("Scripting.Dictionary")
 For Each field In r.Fields: result.Add field.Name, field.Value: Next
 r.Close
 If Nz(result("WorkspaceGuid"), "") <> Nz(MajorScalar("SELECT WorkspaceGuid FROM Workspace WHERE WorkspaceID=1"), "") Then Err.Raise vbObjectError + 9201, "WizardPreviewLoad", "Saved preview belongs to another workspace. Cancel it and preview again."
 Set dom = CreateObject("MSXML2.DOMDocument.6.0"): dom.async = False: dom.resolveExternals = False: dom.setProperty "ProhibitDTD", True
 If Not dom.LoadXML(result("ChoiceXML")) Then Err.Raise vbObjectError + 9202, "WizardPreviewLoad", "Saved preview choices are damaged; no changes were applied."
 Set choices = CreateObject("Scripting.Dictionary")
 For Each node In dom.selectNodes("/choices/choice"): choices.Add CStr(node.getAttribute("key")), CStr(node.Text): Next
 result.Add "Choices", choices
 Set WizardPreviewLoad = result
End Function
Public Function WizardPreviewSave(ByVal kind As String, ByVal path As String, ByVal hash As String, ByVal baseline As String, ByVal baselineHash As String, ByVal report As Long, ByVal revision As Long, ByVal choices As Object, Optional ByVal expected As String = "") As String
 Dim db As DAO.Database, w As DAO.Workspace, r As DAO.Recordset, dom As Object, root As Object, node As Object, key As Variant, id As String, trans As Boolean, failure As Long, message As String, predicate As String
 predicate = WizardWhere(kind)
 If Len(hash) <> 64 Then Err.Raise 5, "WizardPreviewSave", "A source fingerprint is required."
 Set dom = CreateObject("MSXML2.DOMDocument.6.0"): Set root = dom.createElement("choices"): dom.appendChild root
 For Each key In choices.Keys
  Set node = dom.createElement("choice"): node.setAttribute "key", CStr(key): node.Text = CStr(choices(key)): root.appendChild node
 Next
 Set db = CurrentDb: Set w = DBEngine.Workspaces(0)
 On Error GoTo Failed
 w.BeginTrans: trans = True
 Set r = db.OpenRecordset("SELECT StageGuid FROM WizardPreview WHERE " & predicate, dbOpenSnapshot)
 If expected <> "" Then
  If r.EOF Then Err.Raise vbObjectError + 9203, "modWorkspaceImport.WizardPreviewSave", "Preview ownership was cleared. Resume or preview again."
  If r!StageGuid <> expected Then Err.Raise vbObjectError + 9203, "modWorkspaceImport.WizardPreviewSave", "Another preview replaced this one. Resume before making choices."
 End If
 r.Close: Set r = Nothing
 db.Execute "DELETE FROM WizardPreview WHERE " & predicate, dbFailOnError
 id = NewIdentity(): Set r = db.OpenRecordset("WizardPreview", dbOpenDynaset): r.AddNew
 r!StageGuid = id: r!WizardKind = kind
 If WizardPreviewOwner() > 0 Then r!OwnerProfileID = WizardPreviewOwner()
 r!WorkspaceGuid = MajorScalar("SELECT WorkspaceGuid FROM Workspace WHERE WorkspaceID=1")
 r!SourcePath = path: r!SourceHash = hash: r!BaselinePath = baseline: r!BaselineHash = baselineHash
 If report > 0 Then r!TargetReportID = report
 If revision > 0 Then r!ExpectedRevisionID = revision
 r!ChoiceXML = dom.XML: r!SavedAt = Now(): r.Update: r.Close: Set r = Nothing
 w.CommitTrans: trans = False: WizardPreviewSave = id: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modWorkspaceImport", "WizardPreviewSave", "Execute", Err.Number, Err.Description, Err.Source, IIf(trans, "Active", "Not active"))
 failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
 If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
 If trans Then SwiftRollback w, swiftPrimary
 On Error GoTo 0
 SwiftRaise swiftPrimary
End Function
Public Sub WizardPreviewCheck(ByVal kind As String, ByVal token As String)
 Dim saved As Object
 Set saved = WizardPreviewLoad(kind)
 If token = "" Or saved("StageGuid") <> token Then Err.Raise vbObjectError + 9203, "WizardPreviewCheck", "This preview is no longer owned by the current contributor. Resume or preview again."
 If kind = "WorkspaceImport" Then
  If Not saved("Choices").Exists("TargetStamp") Then Err.Raise vbObjectError + 9205, "modWorkspaceImport.WizardPreviewCheck", "Saved workspace preview has no target stamp. Preview again."
  If saved("Choices")("TargetStamp") <> WizardWorkspaceStamp() Then Err.Raise vbObjectError + 9205, "modWorkspaceImport.WizardPreviewCheck", "Workspace setup changed. Cancel or preview again before importing."
 End If
 If SourceFileHash(saved("SourcePath")) <> saved("SourceHash") Then Err.Raise vbObjectError + 9204, "WizardPreviewCheck", "Source changed. Saved choices are retained; cancel or start a new preview."
 If Nz(saved("BaselinePath"), "") <> "" Then
  If SourceFileHash(saved("BaselinePath")) <> saved("BaselineHash") Then Err.Raise vbObjectError + 9204, "WizardPreviewCheck", "Baseline changed. Saved choices are retained; cancel or start a new preview."
 End If
 If Not IsNull(saved("TargetReportID")) Then
  If CurrentRevision(saved("TargetReportID")) <> saved("ExpectedRevisionID") Then Err.Raise vbObjectError + 9205, "WizardPreviewCheck", "Target revision changed. Saved choices are retained; cancel or start a new preview."
 End If
End Sub
Public Function WizardPreviewCancel(ByVal kind As String, Optional ByVal expected As String = "") As Boolean
 Dim db As DAO.Database, predicate As String
 predicate = WizardWhere(kind)
 If expected <> "" Then predicate = predicate & " AND StageGuid=" & QuoteSQL(expected)
 Set db = CurrentDb: db.Execute "DELETE FROM WizardPreview WHERE " & predicate, dbFailOnError
 If expected <> "" And db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 9203, "modWorkspaceImport.WizardPreviewCancel", "The owned preview was replaced; its newer choices were not removed."
 WizardPreviewCancel = True
End Function
Private Function WizardWorkspaceStamp() As String
 Dim rows As DAO.Recordset, field As DAO.Field, value As String, table As Variant
 For Each table In Array("Workspace", "UserProfile")
  Set rows = CurrentDb.OpenRecordset("SELECT * FROM [" & table & "] ORDER BY " & IIf(table = "Workspace", "WorkspaceID", "UserProfileID"), dbOpenSnapshot)
  Do While Not rows.EOF
   For Each field In rows.Fields
    value = Nz(field.Value, "")
    WizardWorkspaceStamp = WizardWorkspaceStamp & Len(value) & ":" & value & ";"
   Next
   rows.MoveNext
  Loop
  rows.Close
 Next
End Function
Public Function WizardImportStagingOnly() As Boolean
 Dim r As DAO.Recordset
 Set r = CurrentDb.OpenRecordset("SELECT TOP 1 StageGuid FROM WizardPreview WHERE NOT (" & WizardWhere("WorkspaceImport") & ")", dbOpenSnapshot)
 WizardImportStagingOnly = r.EOF: r.Close
End Function
'The current app supplies forms/code/schema. Only data is read from the source;
'source VBA/macros are never opened or executed. Migration verifies every field.
Public Function UXChooseSwiftWorkspace() As Boolean
 Dim dialog As Object
 If importing Then Exit Function
 On Error GoTo Failed
 Set dialog = Application.FileDialog(3): dialog.Title = "Import previous Swift NAVFIT workspace": dialog.AllowMultiSelect = False
 dialog.Filters.Clear: dialog.Filters.Add "Swift NAVFIT database", "*.accdb;*.accde"
 If dialog.Show <> -1 Then Exit Function
 UXChooseSwiftWorkspace = UXPreviewSwiftWorkspace(dialog.SelectedItems(1)): Exit Function
Failed:
 MsgBox Err.Description, vbExclamation, "Import Swift NAVFIT data"
End Function
Public Function UXPreviewSwiftWorkspace(ByVal path As String) As Boolean
 Dim reader As DAO.Workspace, source As DAO.Database, r As DAO.Recordset, table As DAO.TableDef, text As String, version As Long, stage As String, failure As Object
 On Error GoTo Failed
 If importing Then Exit Function
 importPath = "": importHash = "": Set workspacePreview = Nothing
 stage = "Open preview source"
 If StrComp(path, CurrentProject.FullName, vbTextCompare) = 0 Then Err.Raise 5, "modWorkspaceImport.UXPreviewSwiftWorkspace", "Choose a different Swift NAVFIT file."
 Set reader = DBEngine.CreateWorkspace("SwiftImportPreview", "admin", "", dbUseJet)
 Set source = reader.OpenDatabase(path, False, True)
 stage = "Read source schema"
 Set r = source.OpenRecordset("SELECT SchemaVersion FROM AppMetadata", dbOpenSnapshot): version = r!SchemaVersion: r.Close: Set r = Nothing
 If version >= 1 And version < 35 Then Err.Raise 5, "modWorkspaceImport.UXPreviewSwiftWorkspace", "This workspace is from before release 0.49.0 (schema " & version & "). Import it into a Swift NAVFIT 0.55.0 Blank first, then import that workspace here. No data was changed."
 If version < 35 Or version > SCHEMA_VERSION Then Err.Raise 5, "modWorkspaceImport.UXPreviewSwiftWorkspace", "This source version is not supported by the current app. No data was changed."
 text = "Source: " & path & vbCrLf & "Schema " & version & " -> " & SCHEMA_VERSION & vbCrLf & vbCrLf
 stage = "Check backup receipt"
 Dim receipt As String
 receipt = BackupReceiptCheck(path, SourceFileHash(path))
 If Len(receipt) > 0 Then text = text & receipt & vbCrLf & vbCrLf
 For Each table In source.TableDefs
  If Left$(table.Name, 4) <> "MSys" And table.Name <> "AppMetadata" Then
   stage = "Count preview table " & table.Name
   Set r = source.OpenRecordset("SELECT Count(*) AS N FROM [" & table.Name & "]", dbOpenSnapshot)
   text = text & table.Name & ": " & r!n & " rows" & vbCrLf: r.Close: Set r = Nothing
  End If
 Next
 stage = "Close preview source"
 source.Close: Set source = Nothing
 reader.Close: Set reader = Nothing
 stage = "Open preview display"
 DoCmd.OpenForm "frmUXWorkspaceImport"
 importHash = SourceFileHash(path): importPath = path
 Dim choices As Object, token As String
 Set choices = CreateObject("Scripting.Dictionary")
 choices.Add "TargetStamp", WizardWorkspaceStamp()
 stage = "Save preview descriptor"
 token = WizardPreviewSave("WorkspaceImport", path, importHash, "", "", 0, 0, choices)
 Set workspacePreview = WizardPreviewLoad("WorkspaceImport")
 stage = "Show preview"
 Forms!frmUXWorkspaceImport!txtPreview = text & vbCrLf & "All workspace data is copied, including revisions, reviews, members, defaults, recovery, history and import provenance. Source code/forms are not imported. Use a new blank app; this does not merge independent workspaces. Initial setup contributor details in an otherwise blank app will be replaced by the source workspace."
 Forms!frmUXWorkspaceImport!cmdApply.Enabled = True
 UXPreviewSwiftWorkspace = True: Exit Function
Failed:
 Set failure = CaptureWorkspaceImportFailure("UXPreviewSwiftWorkspace", stage, Err.Number, Err.Description, Err.Source, "Not started")
 SwiftClose r, failure, "preview recordset"
 SwiftClose source, failure, "preview database"
 SwiftClose reader, failure, "preview workspace"
 On Error Resume Next
 MsgBox "Unable to preview: " & SwiftFailureText(failure), vbExclamation, "Import Swift NAVFIT data"
 If Err.Number <> 0 Then SwiftCleanupError failure, "show preview failure", Err.Number
 On Error GoTo 0
End Function
Public Function ImportSwiftWorkspaceData(ByVal path As String, ByVal hash As String, Optional ByRef transactionOutcome As String = "") As String
 Dim result As String, failure As Long, message As String, source As String, cause As Object
 On Error GoTo Failed
 transactionOutcome = "Not started"
 If StrComp(path, CurrentProject.FullName, vbTextCompare) = 0 Then Err.Raise 5, "modWorkspaceImport.ImportSwiftWorkspaceData", "Cannot import this file into itself."
 WorkspaceImportProgress "Checking source fingerprint and destination..."
 If SourceFileHash(path) <> hash Then Err.Raise 5, "modWorkspaceImport.ImportSwiftWorkspaceData", "Source changed. Choose and preview it again."
 Call BackupReceiptCheck(path, hash)
 result = MigrateWorkspace(path, importing, True, hash)
 If Left$(result, 5) <> "PASS:" Then
  Set cause = SwiftLastFailureContext()
  If Not cause Is Nothing Then
   transactionOutcome = CStr(cause("Transaction"))
   SwiftRaise cause
  End If
  Err.Raise 5, "modWorkspaceImport.ImportSwiftWorkspaceData", result
 End If
 transactionOutcome = "Committed"
 ImportSwiftWorkspaceData = result: Exit Function
Failed:
 failure = Err.Number: message = Err.Description: source = Err.Source
 Err.Raise failure, source, message
End Function
Public Sub WorkspaceImportProgress(ByVal message As String, Optional ByVal forcePaint As Boolean = False)
 'Keep every log entry, but bound expensive control updates to four per second.
 'The modal importer disables actions; guarded pulses process repaint/window events.
 Dim elapsed As Single, stage As String, active As String
 If Not importing Then Exit Sub
 On Error GoTo Failed
 stage = "Build progress text"
 progressText = progressText & Format$(Now, "hh:nn:ss") & "  " & message & vbCrLf
 If Len(progressText) > 28000 Then progressText = Right$(progressText, 24000)
 elapsed = Timer - lastProgressPaint
 If elapsed < 0 Then elapsed = elapsed + 86400!
 If Not forcePaint And elapsed < 0.25! Then Exit Sub
 lastProgressPaint = Timer
 stage = "Paint progress"
 With Forms!frmUXWorkspaceImport
  !txtProgress = progressText
  'No active control is a legitimate optional probe, not an import failure.
  On Error Resume Next
  active = Screen.ActiveControl.Name
  Err.Clear
  On Error GoTo Failed
  If active = "txtProgress" Then !txtProgress.SelStart = Len(progressText)
  !lblStatus.Caption = "Import in progress - " & DateDiff("s", importStarted, Now) & " seconds elapsed. Please wait; keep Access open."
  .Repaint
 End With
PaintStatus:
 stage = "Paint status bar"
 SysCmd acSysCmdSetStatus, message
Pulse:
 WorkspaceImportPulse
 Exit Sub
Failed:
 RecordImportDisplayError displayFailure, stage, Err.Number, Err.Description, Err.Source
 'Painting is advisory. Keep status and event pulses working after a paint fault.
 If stage = "Paint status bar" Then Resume Pulse
 Resume PaintStatus
End Sub
Public Sub WorkspaceImportPulse()
 'Only the modal, action-disabled importer can receive user input during a yield.
 'Do not pump events inside a DAO operation or allow nested progress callbacks.
 If Not importing Or pulseBusy Then Exit Sub
 pulseBusy = True
 On Error GoTo FinishedPulse
 DoEvents
FinishedPulse:
 pulseBusy = False
End Sub
Private Sub LockImportDisplay()
 Dim c As Access.Control, f As Access.Form
 Set f = Forms!frmUXWorkspaceImport
 modalBefore = f.Modal: f.Modal = True
 Set enabledBefore = CreateObject("Scripting.Dictionary")
 For Each c In f.Controls
  If c.ControlType = acCommandButton Then
   enabledBefore.Add c.Name, c.Enabled: c.Enabled = False
  End If
 Next
End Sub
Public Function UXWorkspaceImportBusy() As Boolean
 UXWorkspaceImportBusy = importing
End Function
Private Sub RecordImportDisplayError(ByRef failure As Object, ByVal step As String, ByVal number As Long, ByVal description As String, ByVal source As String, Optional ByVal outcome As String = "Unknown", Optional ByVal procedure As String = "WorkspaceImportProgress")
 If number = 0 Then Exit Sub
 If failure Is Nothing Then Set failure = SwiftFailure("modWorkspaceImport", procedure, step, number, description, source, outcome)
 SwiftCleanupError failure, step, number
End Sub
Private Function CaptureWorkspaceImportFailure(ByVal procedure As String, ByVal stage As String, ByVal number As Long, ByVal description As String, ByVal source As String, ByVal outcome As String) As Object
 Dim cause As Object
 Set cause = SwiftLastFailureContext()
 'Only reuse a context explicitly carried by this error, never stale diagnostics.
 If Not cause Is Nothing Then
  If cause("Number") = number And InStr(description, "Diagnostic: " & cause("Identity")) > 0 Then
   Set CaptureWorkspaceImportFailure = cause: Exit Function
  End If
 End If
 Set CaptureWorkspaceImportFailure = SwiftFailure("modWorkspaceImport", procedure, stage, number, description, source, outcome)
End Function
Private Sub FinishImportDisplay(ByRef failure As Object, ByVal outcome As String)
 Dim name As Variant
 'Release in-memory guards independently of optional UI restoration.
 importing = False: pulseBusy = False
 If Not enabledBefore Is Nothing Then
  For Each name In enabledBefore.Keys
   On Error Resume Next
   Forms!frmUXWorkspaceImport.Controls(CStr(name)).Enabled = enabledBefore(name)
   If Err.Number <> 0 Then RecordImportDisplayError failure, "restore import action " & CStr(name), Err.Number, Err.Description, Err.Source, outcome, "FinishImportDisplay"
   On Error GoTo 0
  Next
 End If
 On Error Resume Next
 Forms!frmUXWorkspaceImport.Modal = modalBefore
 If Err.Number <> 0 Then RecordImportDisplayError failure, "restore import modality", Err.Number, Err.Description, Err.Source, outcome, "FinishImportDisplay"
 On Error GoTo 0
 Set enabledBefore = Nothing
 On Error Resume Next
 DoCmd.Hourglass False
 If Err.Number <> 0 Then RecordImportDisplayError failure, "clear import hourglass", Err.Number, Err.Description, Err.Source, outcome, "FinishImportDisplay"
 On Error GoTo 0
 On Error Resume Next
 SysCmd acSysCmdClearStatus
 If Err.Number <> 0 Then RecordImportDisplayError failure, "clear import status bar", Err.Number, Err.Description, Err.Source, outcome, "FinishImportDisplay"
 On Error GoTo 0
 On Error Resume Next
 Forms!frmUXWorkspaceImport!cmdClose.Enabled = True
 If Err.Number <> 0 Then RecordImportDisplayError failure, "enable import close", Err.Number, Err.Description, Err.Source, outcome, "FinishImportDisplay"
 On Error GoTo 0
 On Error Resume Next
 Forms!frmUXWorkspaceImport.Repaint
 If Err.Number <> 0 Then RecordImportDisplayError failure, "repaint finished import", Err.Number, Err.Description, Err.Source, outcome, "FinishImportDisplay"
 On Error GoTo 0
 If Not failure Is Nothing Then
  failure("Transaction") = outcome
  Call SwiftFailureText(failure)
 End If
End Sub
Public Function UXApplySwiftWorkspace() As Boolean
 Dim result As String, message As String, outcome As String, status As String, safeRetry As Boolean, stage As String
 Dim failure As Object
 If importing Then Exit Function
 Set displayFailure = Nothing
 outcome = "Not started"
 On Error GoTo Failed
 stage = "Validate preview"
 If workspacePreview Is Nothing Then Err.Raise 5, "modWorkspaceImport.UXApplySwiftWorkspace", "Preview ownership was reset. Use Resume saved preview before Apply."
 WizardPreviewCheck "WorkspaceImport", workspacePreview("StageGuid")
 If importPath = "" Then Err.Raise 5, "modWorkspaceImport.UXApplySwiftWorkspace", "Choose and preview a previous Swift file first."
 importStarted = Now: progressText = "": lastProgressPaint = 0: importing = True
 stage = "Lock import display"
 Forms!frmUXWorkspaceImport!txtProgress.SetFocus
 LockImportDisplay
 DoCmd.Hourglass True
 WorkspaceImportProgress "Import started. Copying workspace data, then verifying every original field in database-engine batches.", True
 stage = "Import workspace data"
 result = ImportSwiftWorkspaceData(importPath, importHash, outcome)
 stage = "Show committed import"
 WorkspaceImportProgress "Import committed successfully.", True
 Forms!frmUXWorkspaceImport!lblStatus.Caption = result & ". Close and reopen this app to load the imported workspace."
 importPath = "": importHash = "": Set workspacePreview = Nothing: UXApplySwiftWorkspace = True
 FinishImportDisplay displayFailure, outcome
 Forms!frmUXWorkspaceImport!cmdApply.Enabled = False
 Exit Function
Failed:
 'Capture before any progress painting, control access or cleanup can replace Err.
 Set failure = CaptureWorkspaceImportFailure("UXApplySwiftWorkspace", stage, Err.Number, Err.Description, Err.Source, outcome)
 If Not displayFailure Is Nothing Then
  If Len(displayFailure("Cleanup")) > 0 Then
   If Len(failure("Cleanup")) > 0 Then failure("Cleanup") = failure("Cleanup") & "; "
   failure("Cleanup") = failure("Cleanup") & displayFailure("Cleanup")
  End If
 End If
 Set displayFailure = failure
 message = SwiftFailureText(failure)
 Select Case outcome
 Case "Committed"
  status = "Import data was committed; finalization failed. Close and reopen this workspace before continuing. " & message
  importPath = "": importHash = "": Set workspacePreview = Nothing
 Case "Rolled back"
  status = "Import stopped: " & message & ". Target changes rolled back."
  safeRetry = True
 Case "Not started"
  status = "Import was not started. " & message
  safeRetry = True
 Case Else
  status = "Import stopped; rollback could not be confirmed. Inspect saved data before retrying. " & message
 End Select
 WorkspaceImportProgress status, True
 On Error Resume Next
 Forms!frmUXWorkspaceImport!lblStatus.Caption = status & " Source data is unchanged."
 If Err.Number <> 0 Then SwiftCleanupError failure, "show import failure", Err.Number
 On Error GoTo 0
 FinishImportDisplay displayFailure, outcome
 On Error Resume Next
 Forms!frmUXWorkspaceImport!cmdApply.Enabled = (safeRetry And importPath <> "")
 If Err.Number <> 0 Then SwiftCleanupError failure, "restore import retry", Err.Number
 On Error GoTo 0
 Call SwiftFailureText(failure)
End Function
Public Function UXResumeSwiftWorkspace() As Boolean
 Dim saved As Object
 On Error GoTo Failed
 If importing Then Exit Function
 Set saved = WizardPreviewLoad("WorkspaceImport")
 WizardPreviewCheck "WorkspaceImport", saved("StageGuid")
 UXResumeSwiftWorkspace = UXPreviewSwiftWorkspace(saved("SourcePath")): Exit Function
Failed:
 Forms!frmUXWorkspaceImport!lblStatus.Caption = "Resume unavailable: " & Err.Description
End Function
Public Function UXCancelSwiftWorkspacePreview() As Boolean
 Dim token As String
 If importing Then Exit Function
 On Error GoTo Failed
 If Not workspacePreview Is Nothing Then token = workspacePreview("StageGuid")
 UXCancelSwiftWorkspacePreview = WizardPreviewCancel("WorkspaceImport", token)
 On Error GoTo 0
 importPath = "": importHash = "": Set workspacePreview = Nothing
 Forms!frmUXWorkspaceImport!cmdApply.Enabled = False
 Forms!frmUXWorkspaceImport!lblStatus.Caption = "Saved preview cancelled. No source or workspace data changed."
 Exit Function
Failed:
 Forms!frmUXWorkspaceImport!cmdApply.Enabled = False
 Forms!frmUXWorkspaceImport!lblStatus.Caption = "Cancel unavailable: " & Err.Description & " Saved choices are retained. Resume before cancelling."
End Function
Public Function UXWorkspacePreviewStart() As Boolean
 If importing Then Exit Function
 importPath = "": importHash = "": Set workspacePreview = Nothing
 Forms!frmUXWorkspaceImport.Modal = False
 Forms!frmUXWorkspaceImport!cmdApply.Enabled = False
 Forms!frmUXWorkspaceImport!cmdClose.Enabled = True
 Forms!frmUXWorkspaceImport!lblStatus.Caption = "Choose a source or resume the saved preview. Closing retains it; Cancel removes only this contributor's saved preview."
 UXWorkspacePreviewStart = True
End Function
