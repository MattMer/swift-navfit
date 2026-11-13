Attribute VB_Name = "modDataLink"
Option Compare Database
Option Explicit
' Front end and data file (backlog S01). A blank copy of Swift NAVFIT can write its
' tables into a new data file and then use that file through links. From then on this
' file holds only the application; the workspace is the data file. A missing or
' unusable data file is found at startup and the user is asked where it is.

' Every workspace table in this file, local or linked.
Private Function DataTableNames(ByVal db As DAO.Database) As Collection
    Dim t As DAO.TableDef, names As Collection
    Set names = New Collection
    For Each t In db.TableDefs
        If Left$(t.Name, 4) <> "MSys" And Left$(t.Name, 1) <> "~" Then names.Add t.Name
    Next
    Set DataTableNames = names
End Function

' Relationships between workspace tables; Access keeps a few of its own.
Private Function WorkspaceRelations(ByVal db As DAO.Database) As Long
    Dim relation As DAO.Relation
    For Each relation In db.Relations
        If Left$(relation.Table, 4) <> "MSys" Then WorkspaceRelations = WorkspaceRelations + 1
    Next
End Function

' Nothing has been set up or saved in the local tables.
Private Function LocalWorkspaceBlank() As Boolean
    LocalWorkspaceBlank = (DCount("*", "Workspace") = 0 And DCount("*", "Report") = 0 And DCount("*", "CommandMember") = 0)
End Function

' Forms and modules in a file. A file made by DAO alone has no such containers at all.
Private Function ApplicationObjects(ByVal db As DAO.Database) As Long
    Dim c As DAO.Container
    For Each c In db.Containers
        If c.Name = "Forms" Or c.Name = "Modules" Then ApplicationObjects = ApplicationObjects + c.Documents.Count
    Next
End Function

' "" when the file is a Swift NAVFIT data file this version can use; otherwise why not.
Public Function DataFileProblem(ByVal filePath As String) As String
    Dim db As DAO.Database, r As DAO.Recordset, t As DAO.TableDef, found As Object, name As Variant, version As Long
    On Error GoTo Failed
    If Len(filePath) = 0 Then DataFileProblem = "No data file was chosen.": Exit Function
    If Len(Dir$(filePath)) = 0 Then DataFileProblem = "The data file was not found: " & filePath: Exit Function
    If StrComp(filePath, CurrentProject.FullName, vbTextCompare) = 0 Then DataFileProblem = "That is the Swift NAVFIT application file, not a data file.": Exit Function
    Set db = DBEngine.Workspaces(0).OpenDatabase(filePath, False, True)
    Set found = CreateObject("Scripting.Dictionary"): found.CompareMode = vbTextCompare
    For Each t In db.TableDefs
        found(t.Name) = (Len(t.Connect) = 0)
    Next
    If Not found.Exists("AppMetadata") Or Not found.Exists("Workspace") Or Not found.Exists("Report") Then
        DataFileProblem = "That file is not a Swift NAVFIT data file: " & filePath
        GoTo Done
    End If
    Set r = db.OpenRecordset("SELECT SchemaVersion FROM AppMetadata", dbOpenSnapshot)
    If Not r.EOF Then version = CLng(Nz(r!SchemaVersion, 0))
    r.Close: Set r = Nothing
    'A whole workspace from an earlier version holds the application too; it is imported.
    If ApplicationObjects(db) > 0 Then
        DataFileProblem = "That file is a complete Swift NAVFIT workspace from an earlier version, not a data file. Open it from the setup screen of a newly downloaded Swift NAVFIT, which brings it into a new data file."
        GoTo Done
    End If
    If version > SCHEMA_VERSION Then
        DataFileProblem = "The data file is schema " & version & ", newer than this version of Swift NAVFIT (schema " & SCHEMA_VERSION & "). Use a newer Swift NAVFIT."
        GoTo Done
    End If
    If version <> SCHEMA_VERSION Then
        DataFileProblem = "The data file is schema " & version & " and this version of Swift NAVFIT uses schema " & SCHEMA_VERSION & "." & IIf(version >= UPGRADE_FLOOR, " It needs an upgrade.", " It is too old to upgrade here.")
        GoTo Done
    End If
    For Each name In DataTableNames(CurrentDb)
        If Not found.Exists(CStr(name)) Then
            DataFileProblem = "The data file has no " & name & " table: " & filePath
            GoTo Done
        End If
        If Not found(CStr(name)) Then
            DataFileProblem = "The data file's " & name & " table is itself a link. Choose the file that holds the data."
            GoTo Done
        End If
    Next
Done:
    If Not db Is Nothing Then db.Close
    Exit Function
Failed:
    DataFileProblem = "The data file could not be opened (" & Err.Number & ": " & Err.Description & "): " & filePath
    On Error Resume Next
    If Not r Is Nothing Then r.Close
    If Not db Is Nothing Then db.Close
End Function

' Writes this blank file's tables, with their keys, indexes, rules, relationships and
' built-in rows, into a new data file. Never overwrites, and never moves a workspace
' that already holds data: that goes through the Settings & Help import.
Public Function CreateDataFile(ByVal destination As String) As Boolean
    Dim db As DAO.Database, target As DAO.Database, names As Collection, name As Variant, relation As DAO.Relation, copied As DAO.Relation
    Dim field As DAO.Field, created As Boolean, stage As String, failure As Long, message As String, folder As String, relations As Long
    On Error GoTo Failed
    stage = "Check"
    If SwiftDataIsLinked() Then Err.Raise vbObjectError + 9801, "modDataLink.CreateDataFile", "This copy of Swift NAVFIT already uses a separate data file. Start a new data file from a newly downloaded copy."
    If LCase$(Right$(destination, 6)) <> ".accdb" Then Err.Raise vbObjectError + 9801, "modDataLink.CreateDataFile", "A data file is a new .accdb file."
    folder = Left$(destination, InStrRev(destination, "\") - 1)
    If Len(folder) = 0 Or Len(Dir$(folder, vbDirectory)) = 0 Then Err.Raise vbObjectError + 9801, "modDataLink.CreateDataFile", "The folder for the data file does not exist."
    If Len(Dir$(destination)) > 0 Then Err.Raise vbObjectError + 9801, "modDataLink.CreateDataFile", "That file already exists. A new data file never replaces a file."
    If Not LocalWorkspaceBlank() Then Err.Raise vbObjectError + 9803, "modDataLink.CreateDataFile", "This file already holds a workspace. Move it with Settings & Help > Import data from a previous Swift NAVFIT version."
    Set db = CurrentDb: Set names = DataTableNames(db)
    stage = "Create file"
    Set target = DBEngine.CreateDatabase(destination, dbLangGeneral, dbVersion120): created = True
    target.Close: Set target = Nothing
    stage = "Write tables"
    For Each name In names
        DoCmd.TransferDatabase acExport, "Microsoft Access", destination, acTable, CStr(name), CStr(name), False
    Next
    stage = "Write relationships"
    Set target = DBEngine.Workspaces(0).OpenDatabase(destination)
    For Each relation In db.Relations
        If Left$(relation.Table, 4) <> "MSys" Then
            Set copied = target.CreateRelation(relation.Name, relation.Table, relation.ForeignTable, relation.Attributes)
            For Each field In relation.Fields
                copied.Fields.Append copied.CreateField(field.Name)
                copied.Fields(field.Name).ForeignName = field.ForeignName
            Next
            target.Relations.Append copied
            relations = relations + 1
        End If
    Next
    stage = "Verify"
    For Each name In names
        If target.TableDefs(CStr(name)).Fields.Count <> db.TableDefs(CStr(name)).Fields.Count Then Err.Raise vbObjectError + 9804, "modDataLink.CreateDataFile", "The new data file's " & name & " table differs from this version's."
        If target.TableDefs(CStr(name)).ValidationRule <> db.TableDefs(CStr(name)).ValidationRule Then Err.Raise vbObjectError + 9804, "modDataLink.CreateDataFile", "The new data file's " & name & " rule differs from this version's."
        If target.TableDefs(CStr(name)).RecordCount <> db.TableDefs(CStr(name)).RecordCount Then Err.Raise vbObjectError + 9804, "modDataLink.CreateDataFile", "The new data file's " & name & " rows differ from this version's."
    Next
    If WorkspaceRelations(target) <> relations Then Err.Raise vbObjectError + 9804, "modDataLink.CreateDataFile", "The new data file's relationships differ from this version's."
    target.Close: Set target = Nothing
    CreateDataFile = True
    Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not target Is Nothing Then target.Close
    'Only the file this call created is removed.
    If created Then Kill destination
    On Error GoTo 0
    Err.Raise failure, "modDataLink.CreateDataFile", stage & ": " & message
End Function

' Points every workspace table at the data file. The first time, this file's own blank
' tables are replaced by links; later it repoints the links. Refused when this file
' holds a workspace, so no saved data can be lost here.
Public Function LinkDataFile(ByVal filePath As String) As Boolean
    Dim db As DAO.Database, names As Collection, name As Variant, t As DAO.TableDef, problem As String, stage As String
    Dim failure As Long, message As String, index As Long
    On Error GoTo Failed
    stage = "Check data file"
    problem = DataFileProblem(filePath)
    If Len(problem) > 0 Then Err.Raise vbObjectError + 9802, "modDataLink.LinkDataFile", problem
    Set db = CurrentDb: Set names = DataTableNames(db)
    If Len(db.TableDefs("Workspace").Connect) = 0 Then
        If Not LocalWorkspaceBlank() Then Err.Raise vbObjectError + 9803, "modDataLink.LinkDataFile", "This file already holds a workspace. Move it with Settings & Help > Import data from a previous Swift NAVFIT version."
    End If
    stage = "Remove relationships"
    For index = db.Relations.Count - 1 To 0 Step -1
        If Left$(db.Relations(index).Table, 4) <> "MSys" Then db.Relations.Delete db.Relations(index).Name
    Next
    stage = "Replace tables"
    For Each name In names
        db.TableDefs.Delete CStr(name)
    Next
    stage = "Link tables"
    For Each name In names
        Set t = db.CreateTableDef(CStr(name))
        t.Connect = ";DATABASE=" & filePath: t.SourceTableName = CStr(name)
        db.TableDefs.Append t
    Next
    db.TableDefs.Refresh
    LinkDataFile = True
    Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error GoTo 0
    Err.Raise failure, "modDataLink.LinkDataFile", stage & ": " & message
End Function

' "" when the workspace tables can be used: they are in this file, or they are linked
' to a data file this version can use. Otherwise the reason.
Public Function DataLinkProblem() As String
    If Not SwiftDataIsLinked() Then Exit Function
    DataLinkProblem = DataFileProblem(SwiftDataPath())
End Function

' Startup check. With a missing or unusable data file the user is asked to find it;
' Cancel closes Swift NAVFIT. Under automation there is no one to ask, so it returns
' False and startup stops quietly.
Public Function DataLinkReady() As Boolean
    Dim problem As String, picker As Object, chosen As String
    'An older data file is upgraded first; a failed upgrade leaves it as it was.
    If SwiftDataIsLinked() Then
        problem = UpgradeDataFileIfNeeded(SwiftDataPath())
        If Len(problem) = 0 Then problem = DataLinkProblem() Else problem = "The data file could not be upgraded. " & problem
    End If
    Do While Len(problem) > 0
        If Not Application.UserControl Then Exit Function
        If MsgBox("Swift NAVFIT cannot open its data file." & vbCrLf & vbCrLf & problem & vbCrLf & vbCrLf & "Choose OK to find the data file, or Cancel to close Swift NAVFIT.", vbOKCancel + vbExclamation, "Swift NAVFIT data file") <> vbOK Then
            Application.Quit acQuitSaveNone
            Exit Function
        End If
        Set picker = Application.FileDialog(3)
        picker.Title = "Find your Swift NAVFIT data file": picker.AllowMultiSelect = False
        picker.Filters.Clear: picker.Filters.Add "Swift NAVFIT data file", "*.accdb"
        If picker.Show = -1 Then
            chosen = CStr(picker.SelectedItems(1))
            problem = UpgradeDataFileIfNeeded(chosen)
            If Len(problem) = 0 Then problem = DataFileProblem(chosen) Else problem = "The data file could not be upgraded. " & problem
            If Len(problem) = 0 Then
                LinkDataFile chosen
                problem = DataLinkProblem()
            End If
        End If
    Loop
    SwiftDataHold
    DataLinkReady = True
End Function

' From the setup screen: use a data file. With createNew the file is written first.
' The setup screen is closed, the tables are linked, and startup runs again on the
' data file: setup for a new one, contributor choice for an existing workspace.
Public Function UXUseDataFile(ByVal filePath As String, ByVal createNew As Boolean) As Boolean
    Dim problem As String, screen As Variant
    If createNew Then CreateDataFile filePath
    problem = UpgradeDataFileIfNeeded(filePath)
    If Len(problem) > 0 Then Err.Raise vbObjectError + 9802, "modDataLink.UXUseDataFile", "The data file could not be upgraded. " & problem
    problem = DataFileProblem(filePath)
    If Len(problem) > 0 Then Err.Raise vbObjectError + 9802, "modDataLink.UXUseDataFile", problem
    For Each screen In Array("frmSetup", "frmContributorStart")
        If CurrentProject.AllForms(CStr(screen)).IsLoaded Then DoCmd.Close acForm, CStr(screen), acSaveNo
    Next
    LinkDataFile filePath
    UXUseDataFile = SwiftAppStartup()
End Function

Public Function UXOpenDataFile() As Boolean
    Dim picker As Object
    On Error GoTo Failed
    Set picker = Application.FileDialog(3)
    picker.Title = "Open your Swift NAVFIT data file": picker.AllowMultiSelect = False
    picker.Filters.Clear: picker.Filters.Add "Swift NAVFIT data file", "*.accdb"
    If picker.Show <> -1 Then Exit Function
    'A whole workspace from an earlier version is brought into a new data file instead.
    If IsEarlierWorkspace(CStr(picker.SelectedItems(1))) Then
        UXOpenDataFile = UXBringInWorkspace(CStr(picker.SelectedItems(1)), True)
    Else
        UXOpenDataFile = UXUseDataFile(CStr(picker.SelectedItems(1)), False)
    End If
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Swift NAVFIT data file"
End Function

' True for a complete Swift NAVFIT workspace (tables with the application in one file),
' as every release up to 0.55.0 made.
Public Function IsEarlierWorkspace(ByVal filePath As String) As Boolean
    Dim db As DAO.Database, t As DAO.TableDef, hasMetadata As Boolean
    On Error GoTo Done
    If Len(Dir$(filePath)) = 0 Or SwiftIsWorkspaceFile(filePath) Then Exit Function
    Set db = DBEngine.Workspaces(0).OpenDatabase(filePath, False, True)
    For Each t In db.TableDefs
        If t.Name = "AppMetadata" And Len(t.Connect) = 0 Then hasMetadata = True
    Next
    IsEarlierWorkspace = hasMetadata And ApplicationObjects(db) > 0
Done:
    If Not db Is Nothing Then db.Close
End Function

' From the setup screen of a new copy: a complete workspace from an earlier version is
' brought into a new data file beside it with the verified import, and Swift NAVFIT starts
' on it. The earlier file is only read. With confirm, the user is told first.
Public Function UXBringInWorkspace(ByVal sourcePath As String, ByVal confirm As Boolean) As Boolean
    Dim folder As String, base As String, target As String, number As Long, screen As Variant, message As String
    If SwiftDataIsLinked() Then Err.Raise vbObjectError + 9806, "modDataLink.UXBringInWorkspace", "This copy of Swift NAVFIT already uses a data file. Bring the earlier workspace in from a newly downloaded copy."
    If Not IsEarlierWorkspace(sourcePath) Then Err.Raise vbObjectError + 9806, "modDataLink.UXBringInWorkspace", "That file is not a Swift NAVFIT workspace from an earlier version."
    folder = Left$(sourcePath, InStrRev(sourcePath, "\") - 1)
    base = Mid$(sourcePath, InStrRev(sourcePath, "\") + 1)
    If LCase$(Right$(base, 6)) = ".accdb" Then base = Left$(base, Len(base) - 6)
    target = folder & "\" & base & " data.accdb": number = 1
    Do While Len(Dir$(target)) > 0
        number = number + 1: target = folder & "\" & base & " data " & number & ".accdb"
    Loop
    If confirm And Application.UserControl Then
        If MsgBox("That file is a complete workspace from an earlier version of Swift NAVFIT." & vbCrLf & vbCrLf & "Swift NAVFIT will make a new data file beside it and bring the workspace in, checking every row and field:" & vbCrLf & target & vbCrLf & vbCrLf & "The earlier file is only read, never changed.", vbOKCancel + vbInformation, "Swift NAVFIT data file") <> vbOK Then Exit Function
    End If
    CreateDataFile target
    For Each screen In Array("frmSetup", "frmContributorStart")
        If CurrentProject.AllForms(CStr(screen)).IsLoaded Then DoCmd.Close acForm, CStr(screen), acSaveNo
    Next
    LinkDataFile target
    On Error GoTo ImportFailed
    Call ImportSwiftWorkspaceData(sourcePath, SourceFileHash(sourcePath))
    On Error GoTo 0
    UXBringInWorkspace = SwiftAppStartup()
    Exit Function
ImportFailed:
    message = Err.Description
    On Error Resume Next
    Call SwiftAppStartup
    On Error GoTo 0
    Err.Raise vbObjectError + 9807, "modDataLink.UXBringInWorkspace", "The workspace could not be brought in: " & message & " The earlier file was not changed. This copy now uses a new, empty data file (" & target & "); set it up, or try Settings & Help > Import data from a previous Swift NAVFIT version again."
End Function

Public Function UXNewDataFile() As Boolean
    Dim picker As Object, folder As String, filePath As String, number As Long
    On Error GoTo Failed
    Set picker = Application.FileDialog(4)
    picker.Title = "Choose a folder for the new Swift NAVFIT data file"
    If picker.Show <> -1 Then Exit Function
    folder = CStr(picker.SelectedItems(1))
    filePath = folder & "\Swift NAVFIT data.accdb": number = 1
    Do While Len(Dir$(filePath)) > 0
        number = number + 1: filePath = folder & "\Swift NAVFIT data " & number & ".accdb"
    Loop
    UXNewDataFile = UXUseDataFile(filePath, True)
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Swift NAVFIT data file"
End Function
