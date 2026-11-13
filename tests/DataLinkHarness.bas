Attribute VB_Name = "DataLinkHarness"
Option Explicit
'Front end and data file (S01): create a data file from a blank copy, link to it,
'work through the links, and detect a data file that is missing or unusable.
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 9990, , message
    checks = checks + 1
End Sub
Private Function CreateRefused(ByVal destination As String) As Boolean
    On Error GoTo Failed
    Call CreateDataFile(destination)
    Exit Function
Failed:
    CreateRefused = True
End Function
Private Function LinkRefused(ByVal filePath As String) As Boolean
    On Error GoTo Failed
    Call LinkDataFile(filePath)
    Exit Function
Failed:
    LinkRefused = True
End Function
Private Function LinkedTables() As Long
    Dim t As DAO.TableDef
    For Each t In CurrentDb.TableDefs
        If Left$(t.Name, 4) <> "MSys" And Left$(t.Name, 1) <> "~" And Len(t.Connect) > 0 Then LinkedTables = LinkedTables + 1
    Next
End Function
Private Function LocalTables() As Long
    Dim t As DAO.TableDef
    For Each t In CurrentDb.TableDefs
        If Left$(t.Name, 4) <> "MSys" And Left$(t.Name, 1) <> "~" And Len(t.Connect) = 0 Then LocalTables = LocalTables + 1
    Next
End Function

'On a blank single file: nothing is linked, and a data file is written beside it.
Public Function TestDataFileCreate(ByVal dataPath As String) As String
    Dim tables As Long
    On Error GoTo Failed
    checks = 0
    tables = LocalTables()
    Check tables > 80 And LinkedTables() = 0, "A build starts with local tables"
    Check Not SwiftDataIsLinked() And SwiftDataPath() = CurrentProject.FullName, "A single file is its own data"
    Check DataLinkProblem() = "" And DataLinkReady(), "A single file needs no data file"
    Check DataFileProblem(dataPath) Like "The data file was not found*", "A missing file is named: " & DataFileProblem(dataPath)
    Check DataFileProblem(CurrentProject.FullName) Like "That is the Swift NAVFIT application file*", "The application file is not a data file"
    Check CreateRefused(Left$(dataPath, Len(dataPath) - 6) & ".txt"), "A data file must be an .accdb"
    Check CreateDataFile(dataPath), "Data file was not created"
    Check Len(Dir$(dataPath)) > 0 And DataFileProblem(dataPath) = "", "New data file is not usable: " & DataFileProblem(dataPath)
    Check CreateRefused(dataPath), "An existing file was overwritten"
    Check LocalTables() = tables And LinkedTables() = 0, "Creating a data file changed this file"
    TestDataFileCreate = "PASS: " & checks & " data file creation assertions; " & tables & " tables": Exit Function
Failed:
    TestDataFileCreate = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function

'Link to the data file, then set up and save through the links.
Public Function TestDataFileLink(ByVal dataPath As String, ByVal otherPath As String) As String
    Dim tables As Long, id As Long
    On Error GoTo Failed
    checks = 0
    tables = LocalTables()
    Check LinkRefused(otherPath), "A file that is not a data file was linked"
    Check LocalTables() = tables, "A refused link changed this file"
    Check LinkDataFile(dataPath), "Link failed"
    Check LinkedTables() = tables And LocalTables() = 0, "Not every table is linked: " & LinkedTables() & " of " & tables
    Check SwiftDataIsLinked() And StrComp(SwiftDataPath(), dataPath, vbTextCompare) = 0, "Data path is " & SwiftDataPath()
    Check SwiftIsWorkspaceFile(dataPath) And SwiftIsWorkspaceFile(CurrentProject.FullName), "Both files are the workspace"
    Check DataLinkProblem() = "" And DataLinkReady(), "Linked data file not ready: " & DataLinkProblem()
    Check CreateRefused(Left$(dataPath, Len(dataPath) - 6) & "-second.accdb"), "A linked copy created a data file"
    Check Nz(DLookup("SchemaVersion", "AppMetadata"), 0) = SCHEMA_VERSION, "Schema is not readable through the link"
    Check ConfigureWorkspace("SYNTHETIC DATA FILE COMMAND", "Tester"), "Setup failed through the link"
    id = CommitDraft(0, "EVAL", "Linked draft", "DOE, PAT", #1/1/2026#, #9/30/2026#, "", "Data file test")
    Check id > 0 And DCount("*", "Report") = 1, "Draft did not save through the link"
    Check DCount("*", "RevisionValue", "RevisionID=" & CurrentRevision(id)) > 0, "Revision values missing"
    Check LinkDataFile(dataPath), "Relink to the same file failed"
    Check DCount("*", "Report") = 1 And LinkedTables() = tables, "Relink lost the link or the data"
    TestDataFileLink = "PASS: " & checks & " link, setup and save assertions": Exit Function
Failed:
    TestDataFileLink = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function

'A new session on the linked file: the data is still there. Then the data file is
'reported missing, and pointing at the moved copy restores the workspace.
Public Function TestDataFileMoved(ByVal oldPath As String, ByVal movedPath As String) As String
    On Error GoTo Failed
    checks = 0
    Check SwiftDataIsLinked() And StrComp(SwiftDataPath(), oldPath, vbTextCompare) = 0, "Link was not kept: " & SwiftDataPath()
    Check DataLinkProblem() Like "The data file was not found*", "Missing data file not reported: " & DataLinkProblem()
    Check Not DataLinkReady(), "Startup went on without its data file"
    Check LinkDataFile(movedPath), "Relink to the moved file failed"
    Check DataLinkProblem() = "" And DataLinkReady(), "Moved data file not ready: " & DataLinkProblem()
    Check DCount("*", "Report") = 1 And Nz(DLookup("CommandName", "Workspace"), "") = "SYNTHETIC DATA FILE COMMAND", "Saved data is not in the moved file"
    TestDataFileMoved = "PASS: " & checks & " missing data file and relink assertions": Exit Function
Failed:
    TestDataFileMoved = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function

'A verified backup written while the tables are in a separate data file.
Public Function TestDataFileBackup(ByVal destination As String) As String
    Dim evidence As String
    On Error GoTo Failed
    checks = 0
    Check SwiftDataIsLinked(), "Backup test needs a linked copy"
    evidence = WriteWorkspaceBackup(destination)
    Check InStr(evidence, """Status"":""PASS""") > 0 And InStr(evidence, """RowsVerified"":") > 0, "Backup evidence: " & Left$(evidence, 200)
    Check InStr(1, evidence, Replace(SwiftDataPath(), "\", "\\"), vbTextCompare) > 0, "Backup source is not the data file: " & Left$(evidence, 200)
    Check Len(Dir$(destination)) > 0 And Len(Dir$(destination & ".verification.json")) > 0, "Backup files missing"
    Check DataFileProblem(destination) = "", "Backup is not a usable data file: " & DataFileProblem(destination)
    Check DCount("*", "Report") = 1 And SwiftDataIsLinked(), "Backup changed the workspace"
    TestDataFileBackup = "PASS: " & checks & " linked backup assertions": Exit Function
Failed:
    TestDataFileBackup = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function

'Fill in and save the open setup screen, as its Create workspace button does.
Public Function SaveSetupValues(ByVal commandName As String, ByVal contributor As String) As Boolean
    Forms!frmSetup!txtCommand = commandName: Forms!frmSetup!txtProfile = contributor
    SaveSetupValues = SaveSetup() And DCount("*", "Workspace") = 1 And CurrentProject.AllForms("frmHome").IsLoaded
End Function
