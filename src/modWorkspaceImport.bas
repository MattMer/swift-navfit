Attribute VB_Name = "modWorkspaceImport"
Option Explicit
Private importPath As String, importHash As String
Private importing As Boolean, importStarted As Date, progressText As String, lastProgressPaint As Single
Private pulseBusy As Boolean, enabledBefore As Object, modalBefore As Boolean
'The current app supplies forms/code/schema. Only data is read from the source;
'source VBA/macros are never opened or executed. Migration verifies every field.
Public Function UXChooseSwiftWorkspace() As Boolean
 Dim dialog As Object
 If importing Then Exit Function
 On Error GoTo Failed
 Set dialog = Application.FileDialog(3): dialog.Title = "Import previous Swift NAVFIT workspace": dialog.AllowMultiSelect = False
 dialog.Filters.Clear: dialog.Filters.Add "Swift NAVFIT database", "*.accdb"
 If dialog.Show <> -1 Then Exit Function
 UXChooseSwiftWorkspace = UXPreviewSwiftWorkspace(dialog.SelectedItems(1)): Exit Function
Failed:
 MsgBox Err.Description, vbExclamation, "Import Swift NAVFIT data"
End Function
Public Function UXPreviewSwiftWorkspace(ByVal path As String) As Boolean
 Dim reader As DAO.Workspace, source As DAO.Database, r As DAO.Recordset, table As DAO.TableDef, text As String, version As Long, message As String
 On Error GoTo Failed
 If importing Then Exit Function
 importPath = "": importHash = ""
 If StrComp(path, CurrentProject.FullName, vbTextCompare) = 0 Then Err.Raise 5, , "Choose a different Swift NAVFIT file."
 Set reader = DBEngine.CreateWorkspace("SwiftImportPreview", "admin", "", dbUseJet)
 Set source = reader.OpenDatabase(path, False, True)
 Set r = source.OpenRecordset("SELECT SchemaVersion FROM AppMetadata", dbOpenSnapshot): version = r!SchemaVersion: r.Close
 If version < 1 Or version > SCHEMA_VERSION Then Err.Raise 5, , "This source version is not supported by the current app. No data was changed."
 text = "Source: " & path & vbCrLf & "Schema " & version & " -> " & SCHEMA_VERSION & vbCrLf & vbCrLf
 For Each table In source.TableDefs
  If Left$(table.Name, 4) <> "MSys" And table.Name <> "AppMetadata" Then
   Set r = source.OpenRecordset("SELECT Count(*) AS N FROM [" & table.Name & "]", dbOpenSnapshot)
   text = text & table.Name & ": " & r!n & " rows" & vbCrLf: r.Close
  End If
 Next
 source.Close: reader.Close
 importHash = SourceFileHash(path): importPath = path
 DoCmd.OpenForm "frmUXWorkspaceImport"
 Forms!frmUXWorkspaceImport!txtPreview = text & vbCrLf & "All workspace data is copied, including revisions, reviews, members, defaults, recovery, history and import provenance. Source code/forms are not imported. Use a new blank app; this does not merge independent workspaces. Initial setup contributor details in an otherwise blank app will be replaced by the source workspace."
 Forms!frmUXWorkspaceImport!cmdApply.Enabled = True
 UXPreviewSwiftWorkspace = True: Exit Function
Failed:
 message = Err.Description
 On Error Resume Next: If Not source Is Nothing Then source.Close
 If Not reader Is Nothing Then reader.Close
 MsgBox "Unable to preview: " & message, vbExclamation, "Import Swift NAVFIT data"
End Function
Public Function ImportSwiftWorkspaceData(ByVal path As String, ByVal hash As String) As String
 Dim result As String, failure As Long, message As String
 On Error GoTo Failed
 If StrComp(path, CurrentProject.FullName, vbTextCompare) = 0 Then Err.Raise 5, , "Cannot import this file into itself."
 WorkspaceImportProgress "Checking source fingerprint and destination..."
 If SourceFileHash(path) <> hash Then Err.Raise 5, , "Source changed. Choose and preview it again."
 result = MigrateWorkspace(path, importing, True, hash)
 If Left$(result, 5) <> "PASS:" Then Err.Raise 5, , result
 ImportSwiftWorkspaceData = result: Exit Function
Failed:
 failure = Err.Number: message = Err.Description
 Err.Raise failure, "ImportSwiftWorkspaceData", message
End Function
Public Sub WorkspaceImportProgress(ByVal message As String, Optional ByVal forcePaint As Boolean = False)
 'Keep every log entry, but bound expensive control updates to four per second.
 'The modal importer disables actions; guarded pulses process repaint/window events.
 Dim elapsed As Single
 If Not importing Then Exit Sub
 On Error Resume Next
 progressText = progressText & Format$(Now, "hh:nn:ss") & "  " & message & vbCrLf
 If Len(progressText) > 28000 Then progressText = Right$(progressText, 24000)
 elapsed = Timer - lastProgressPaint
 If elapsed < 0 Then elapsed = elapsed + 86400!
 If Not forcePaint And elapsed < 0.25! Then Exit Sub
 lastProgressPaint = Timer
 With Forms!frmUXWorkspaceImport
  !txtProgress = progressText
  If Screen.ActiveControl.Name = "txtProgress" Then !txtProgress.SelStart = Len(progressText)
  !lblStatus.Caption = "Import in progress - " & DateDiff("s", importStarted, Now) & " seconds elapsed. Please wait; keep Access open."
  .Repaint
 End With
 SysCmd acSysCmdSetStatus, message
 WorkspaceImportPulse
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
Private Sub FinishImportDisplay()
 On Error Resume Next
 Dim name As Variant
 If Not enabledBefore Is Nothing Then
  For Each name In enabledBefore.Keys
   Forms!frmUXWorkspaceImport.Controls(CStr(name)).Enabled = enabledBefore(name)
  Next
 End If
 Forms!frmUXWorkspaceImport.Modal = modalBefore
 importing = False
 Set enabledBefore = Nothing
 DoCmd.Hourglass False
 SysCmd acSysCmdClearStatus
 Forms!frmUXWorkspaceImport!cmdClose.Enabled = True
 Forms!frmUXWorkspaceImport.Repaint
End Sub
Public Function UXApplySwiftWorkspace() As Boolean
 Dim result As String, message As String
 If importing Then Exit Function
 On Error GoTo Failed
 If importPath = "" Then Err.Raise 5, , "Choose and preview a previous Swift file first."
 importStarted = Now: progressText = "": lastProgressPaint = 0: importing = True
 Forms!frmUXWorkspaceImport!txtProgress.SetFocus
 LockImportDisplay
 DoCmd.Hourglass True
 WorkspaceImportProgress "Import started. Copying workspace data, then verifying every original field in database-engine batches.", True
 result = ImportSwiftWorkspaceData(importPath, importHash)
 WorkspaceImportProgress "Import committed successfully.", True
 Forms!frmUXWorkspaceImport!lblStatus.Caption = result & ". Close and reopen this app to load the imported workspace."
 importPath = "": importHash = "": UXApplySwiftWorkspace = True
 FinishImportDisplay
 Forms!frmUXWorkspaceImport!cmdApply.Enabled = False
 Exit Function
Failed:
 message = Err.Description
 WorkspaceImportProgress "Import stopped: " & message & ". Target changes rolled back.", True
 Forms!frmUXWorkspaceImport!lblStatus.Caption = message & " Source data is unchanged."
 FinishImportDisplay
 Forms!frmUXWorkspaceImport!cmdApply.Enabled = (importPath <> "")
End Function
