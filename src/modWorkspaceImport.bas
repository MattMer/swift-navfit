Attribute VB_Name = "modWorkspaceImport"
Option Explicit
Private importPath As String, importHash As String
'The current app supplies forms/code/schema. Only data is read from the source;
'source VBA/macros are never opened or executed. Migration verifies every field.
Public Function UXChooseSwiftWorkspace() As Boolean
 Dim dialog As Object
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
 Dim db As DAO.Database, table As DAO.TableDef, r As DAO.Recordset, w As DAO.Workspace, result As String, trans As Boolean, failure As Long, message As String
 On Error GoTo Failed
 If StrComp(path, CurrentProject.FullName, vbTextCompare) = 0 Then Err.Raise 5, , "Cannot import this file into itself."
 If SourceFileHash(path) <> hash Then Err.Raise 5, , "Source changed. Choose and preview it again."
 Set db = CurrentDb
 For Each table In db.TableDefs
  If Left$(table.Name, 4) <> "MSys" And table.Name <> "AppMetadata" And table.Name <> "Workspace" And table.Name <> "UserProfile" Then
   Set r = db.OpenRecordset("SELECT Count(*) AS N FROM [" & table.Name & "]", dbOpenSnapshot)
   If r!n > 0 Then r.Close: Err.Raise 5, , "This app already contains workspace data. Open a new blank release and import there; existing data will not be overwritten."
   r.Close
  End If
 Next
 Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
 db.Execute "DELETE FROM Workspace", dbFailOnError: db.Execute "DELETE FROM UserProfile", dbFailOnError
 result = MigrateWorkspace(path)
 If Left$(result, 5) <> "PASS:" Then Err.Raise 5, , result
 If SourceFileHash(path) <> hash Then Err.Raise 5, , "Source changed during transfer. No data imported."
 w.CommitTrans: trans = False
 ImportSwiftWorkspaceData = result: Exit Function
Failed:
 failure = Err.Number: message = Err.Description
 On Error Resume Next: If trans Then w.Rollback
 On Error GoTo 0: Err.Raise failure, "ImportSwiftWorkspaceData", message
End Function
Public Function UXApplySwiftWorkspace() As Boolean
 Dim result As String
 On Error GoTo Failed
 If importPath = "" Then Err.Raise 5, , "Choose and preview a previous Swift file first."
 result = ImportSwiftWorkspaceData(importPath, importHash)
 Forms!frmUXWorkspaceImport!cmdApply.Enabled = False
 Forms!frmUXWorkspaceImport!lblStatus.Caption = result & ". Close and reopen this app to load the imported workspace."
 importPath = "": importHash = "": UXApplySwiftWorkspace = True: Exit Function
Failed:
 Forms!frmUXWorkspaceImport!lblStatus.Caption = Err.Description & " Source data is unchanged."
End Function
