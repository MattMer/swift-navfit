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
        If SwiftWorkspaceTable(t.Name) And Len(t.Connect) > 0 Then LinkedTables = LinkedTables + 1
    Next
End Function
Private Function LocalTables() As Long
    Dim t As DAO.TableDef
    For Each t In CurrentDb.TableDefs
        If SwiftWorkspaceTable(t.Name) And Len(t.Connect) = 0 Then LocalTables = LocalTables + 1
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
Private Function BlankTables() As Long
    Dim t As DAO.TableDef
    For Each t In CurrentDb.TableDefs
        If Left$(t.Name, 9) = "USysBlank" Then BlankTables = BlankTables + 1
    Next
End Function

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
    Check BlankTables() = tables, "The blank tables were not kept for a new data file: " & BlankTables()
    Check CanCreateDataFile() And CreateDataFile(Left$(dataPath, Len(dataPath) - 6) & "-fresh.accdb"), "A linked copy could not make a new data file"
    Check DataFileProblem(Left$(dataPath, Len(dataPath) - 6) & "-fresh.accdb") = "", "The new data file is not usable"
    Check StrComp(SwiftDataPath(), dataPath, vbTextCompare) = 0, "Making a new data file changed the link"
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

'Other sessions are counted on the data file, not on this application file.
Public Function TestDataFileSessions(ByVal expected As Long) As String
    Dim computers As String, others As Long, notice As String
    On Error GoTo Failed
    checks = 0
    Check SwiftDataIsLinked(), "Session test needs a linked copy"
    Check DCount("*", "Report") >= 0, "Data file is not readable"
    others = OtherWorkspaceSessions(computers)
    Check others = expected, "Counted " & others & " other sessions (" & computers & "), expected " & expected
    notice = ConcurrentSessionWarning(others, computers)
    If expected = 0 Then
        Check computers = "" And notice = "", "A lone session was warned: " & notice
    Else
        Check notice Like "This workspace is already open in another session*", "Warning: " & notice
    End If
    TestDataFileSessions = "PASS: " & checks & " session assertions; " & others & " other session(s) on the data file": Exit Function
Failed:
    TestDataFileSessions = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function

'Fill in and save the open setup screen, as its Create workspace button does.
Public Function SaveSetupValues(ByVal commandName As String, ByVal contributor As String) As Boolean
    Forms!frmSetup!txtCommand = commandName: Forms!frmSetup!txtProfile = contributor
    SaveSetupValues = SaveSetup() And DCount("*", "Workspace") = 1 And CurrentProject.AllForms("frmHome").IsLoaded
End Function

'Upgrade a data file: "PASS: " and the path of the copy from before, or why not.
Public Function TryUpgrade(ByVal filePath As String) As String
    On Error GoTo Failed
    TryUpgrade = "PASS: " & UpgradeDataFile(filePath): Exit Function
Failed:
    TryUpgrade = "REJECTED: " & Err.Number & " " & Err.Description
End Function

Public Function UpgradeFloorValue() As Long
    UpgradeFloorValue = UPGRADE_FLOOR
End Function

'Create workspace on the setup screen, as its button does, and report where the data went.
Public Function TestSetupDefault(ByVal commandName As String, ByVal contributor As String) As String
    Dim expected As String
    On Error GoTo Failed
    checks = 0
    expected = CurrentProject.Path & "\Swift NAVFIT data.accdb"
    Check CurrentProject.AllForms("frmSetup").IsLoaded And Not SwiftDataIsLinked(), "Setup is not open on an unlinked copy"
    Check Forms!frmSetup!uxDataFileLocation.Caption Like "Your work will be kept in a new data file, " & expected & "*", "Setup line: " & Forms!frmSetup!uxDataFileLocation.Caption
    Check Forms!frmSetup!cmdNewDataFile.Enabled, "Change folder is not available before setup"
    'A missing name is refused before any data file is made (the message box is not opened here).
    Forms!frmSetup!txtCommand = commandName: Forms!frmSetup!txtProfile = contributor: Forms!frmSetup!txtCommandUIC = "n1234"
    Check UXCreateWorkspace(), "Create workspace failed"
    Check SwiftDataIsLinked() And StrComp(SwiftDataPath(), expected, vbTextCompare) = 0, "Data file is " & SwiftDataPath()
    Check DCount("*", "Workspace") = 1 And Nz(DLookup("CommandName", "Workspace"), "") = commandName, "Workspace not set up in the data file"
    Check CurrentProject.AllForms("frmHome").IsLoaded, "Overview did not open"
    Check Nz(DLookup("CommandUIC", "Workspace"), "") = "N1234" And CommandHeading() = commandName & " (N1234)", "Setup UIC not stored: " & CommandHeading()
    TestSetupDefault = "PASS: " & checks & " default data file setup assertions": Exit Function
Failed:
    TestSetupDefault = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function

'The window shown when the data file cannot be opened, and its Create a new data file path.
Public Function TestDataFileWindow(ByVal newPath As String) As String
    Dim problem As String
    On Error GoTo Failed
    checks = 0
    problem = DataLinkProblem()
    Check problem Like "The data file was not found*", "Missing data file not reported: " & problem
    DoCmd.OpenForm "frmDataFileMissing", , , , , , problem
    Check Forms!frmDataFileMissing!lblProblem.Caption = problem, "The window does not say why"
    Check Forms!frmDataFileMissing!cmdFindDataFile.Caption = "Find the data file..." And Forms!frmDataFileMissing!cmdCloseSwift.Caption = "Close Swift NAVFIT", "Window buttons"
    Check Forms!frmDataFileMissing!cmdCreateDataFile.Enabled, "Create a new data file is not offered"
    Check UXDataFileChoice("new") And Not CurrentProject.AllForms("frmDataFileMissing").IsLoaded, "The window did not close on a choice"
    Check CreateAndUseDataFile(newPath) = "", "A new data file could not be made and used"
    Check StrComp(SwiftDataPath(), newPath, vbTextCompare) = 0 And DataLinkProblem() = "", "Not linked to the new data file"
    Check DCount("*", "Workspace") = 0 And DCount("*", "Report") = 0, "The new data file is not empty"
    TestDataFileWindow = "PASS: " & checks & " missing data file window and new data file assertions": Exit Function
Failed:
    TestDataFileWindow = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function

'Presence helpers for the two-session test (PresenceEditing and PresenceEnd are subs).
Public Function PresenceEditingFor(ByVal reportID As Long) As Boolean
    PresenceEditing reportID
    PresenceEditingFor = True
End Function

Public Function PresenceEndFor() As Boolean
    PresenceEnd
    PresenceEndFor = True
End Function

Private Function UICRefused(ByVal value As String) As Boolean
    On Error GoTo Failed
    Call CommandUICText(value)
    Exit Function
Failed:
    UICRefused = True
End Function

'Settings & Help > Command name and UIC, and a new report starting with the command's UIC.
Public Function TestCommandDetails() As String
    Dim journal As Long
    On Error GoTo Failed
    checks = 0
    Check UICRefused("ABC-1") And UICRefused("123456") And Not UICRefused(" ab12 "), "UIC rules"
    journal = DCount("*", "OperationJournal")
    Check CommandDetailsSave("SYNTHETIC RENAMED COMMAND", "r5432"), "Saving command details failed"
    Check CommandHeading() = "SYNTHETIC RENAMED COMMAND (R5432)", "Heading: " & CommandHeading()
    Check DCount("*", "OperationJournal") = journal + 1, "The change is not in operation history"
    DoCmd.OpenForm "frmEditor", , , , , , "0"
    Check Nz(Forms!frmEditor!txtUnitUIC, "") = "R5432", "A new report did not start with the command UIC: " & Nz(Forms!frmEditor!txtUnitUIC, "")
    DoCmd.Close acForm, "frmEditor", acSaveNo
    Check CommandDetailsSave("SYNTHETIC RENAMED COMMAND", "") And IsNull(DLookup("CommandUIC", "Workspace")), "A blank UIC was not cleared"
    DoCmd.OpenForm "frmUXCommandDetails"
    Check Forms!frmUXCommandDetails!txtDetailsCommand = "SYNTHETIC RENAMED COMMAND", "The window does not show the command name"
    Check UXCommandDetailsClose(), "The window did not close"
    TestCommandDetails = "PASS: " & checks & " command name and UIC assertions": Exit Function
Failed:
    TestCommandDetails = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
