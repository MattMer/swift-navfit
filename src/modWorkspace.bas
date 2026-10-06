Attribute VB_Name = "modWorkspace"
Option Compare Database
Option Explicit

Public Function CreateFolder(ByVal folderName As String, Optional ByVal parentID As Long = 0) As Long
    Dim db As DAO.Database, r As DAO.Recordset
    folderName = Trim$(folderName)
    If Len(folderName) = 0 Or Len(folderName) > 80 Then Err.Raise vbObjectError + 200, , "Enter a folder name of 1 to 80 characters."
    If parentID <> 0 And DCount("*", "Folder", "FolderID=" & parentID) = 0 Then Err.Raise vbObjectError + 201, , "Parent folder does not exist."
    Set db = CurrentDb
    Set r = db.OpenRecordset("Folder", dbOpenDynaset)
    r.AddNew: r!FolderGuid = NewIdentity(): r!FolderName = folderName
    If parentID <> 0 Then r!ParentFolderID = parentID
    r.Update: r.Bookmark = r.LastModified: CreateFolder = r!FolderID: r.Close
End Function

Public Function FolderPath(ByVal folderID As Long) As String
    Dim parent As Long, depth As Long, part As String
    Do While folderID > 0
        If depth > 40 Then Err.Raise vbObjectError + 202, , "Folder nesting exceeds supported depth."
        part = Nz(DLookup("FolderName", "Folder", "FolderID=" & folderID), "")
        If Len(part) = 0 Then Err.Raise vbObjectError + 203, , "Folder path is incomplete."
        If Len(FolderPath) > 0 Then FolderPath = part & " / " & FolderPath Else FolderPath = part
        parent = Nz(DLookup("ParentFolderID", "Folder", "FolderID=" & folderID), 0)
        folderID = parent: depth = depth + 1
    Loop
End Function

Public Function MoveFolder(ByVal folderID As Long, ByVal parentID As Long) As Boolean
    Dim cursor As Long, depth As Long, q As DAO.QueryDef
    If DCount("*", "Folder", "FolderID=" & folderID) = 0 Then Err.Raise vbObjectError + 204, , "Folder does not exist."
    cursor = parentID
    Do While cursor <> 0
        If cursor = folderID Then Err.Raise vbObjectError + 205, , "A folder cannot be moved into itself or its descendants."
        If DCount("*", "Folder", "FolderID=" & cursor) = 0 Then Err.Raise vbObjectError + 206, , "Destination folder does not exist."
        cursor = Nz(DLookup("ParentFolderID", "Folder", "FolderID=" & cursor), 0)
        depth = depth + 1
        If depth > 40 Then Err.Raise vbObjectError + 207, , "Folder path is too deep or contains a cycle."
    Loop
    Set q = CurrentDb.CreateQueryDef("", "PARAMETERS parent Long, id Long; UPDATE Folder SET ParentFolderID=parent WHERE FolderID=id")
    If parentID = 0 Then q!parent = Null Else q!parent = parentID
    q!id = folderID: q.Execute dbFailOnError: q.Close
    MoveFolder = True
End Function

Public Function AssignFolder(ByVal reportID As Long, ByVal folderID As Long) As Boolean
    Dim w As DAO.Workspace, db As DAO.Database, q As DAO.QueryDef, r As DAO.Recordset
    Dim oldFolder As Variant, pid As Long, failure As Long, message As String
    pid = Nz(DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1"), 0)
    If pid = 0 Then Err.Raise vbObjectError + 208, , "Set up a contributor profile first."
    If folderID <> 0 And DCount("*", "Folder", "FolderID=" & folderID) = 0 Then Err.Raise vbObjectError + 209, , "Destination folder does not exist."
    Set w = DBEngine.Workspaces(0): Set db = CurrentDb
    On Error GoTo Failed
    w.BeginTrans
    Set r = db.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 210, , "Report does not exist."
    oldFolder = r!FolderID
    If r!IsArchived Then Err.Raise vbObjectError + 211, , "Recover the archived report before moving it."
    r.Close
    If Nz(oldFolder, 0) = folderID Then w.Rollback: AssignFolder = True: Exit Function
    Set q = db.CreateQueryDef("", "PARAMETERS folder Long, id Long; UPDATE Report SET FolderID=folder WHERE ReportID=id")
    If folderID = 0 Then q!folder = Null Else q!folder = folderID
    q!id = reportID: q.Execute dbFailOnError: q.Close
    Set r = db.OpenRecordset("FolderMove", dbOpenDynaset)
    r.AddNew: r!ReportID = reportID: r!OldFolderID = oldFolder
    If folderID <> 0 Then r!NewFolderID = folderID
    r!AuthorProfileID = pid: r!MovedAt = Now(): r.Update: r.Close
    w.CommitTrans: AssignFolder = True: Exit Function
Failed:
    failure = Err.Number: message = Err.Description: w.Rollback
    Err.Raise failure, "AssignFolder", message
End Function

Public Function RefreshFolders() As Boolean
    Dim f As Form
    Set f = Forms!frmLibrary
    If UXHasControl(f, "uxGroupOrganization") Then RefreshFolders = UXRefreshGroupFilter(): Exit Function
    f!lstFolders.RowSource = "SELECT 0 AS ID, 'All reports' AS Label FROM AppMetadata UNION ALL SELECT -1, 'Unassigned' FROM AppMetadata UNION ALL SELECT FolderID, FolderPath([FolderID]) FROM Folder ORDER BY Label"
    f!lstFolders.Requery
    f!cboDestination.RowSource = "SELECT 0 AS ID, 'Unassigned' AS Label FROM AppMetadata UNION ALL SELECT FolderID, FolderPath([FolderID]) FROM Folder ORDER BY Label"
    f!cboDestination.Requery
    If IsNull(f!lstFolders) Then f!lstFolders = 0
    If UXHasControl(f, "uxFolderFilter") Then
        f!uxFolderFilter.Requery
        f!uxFolderFilter = f!lstFolders
    End If
    If IsNull(f!cboDestination) Then f!cboDestination = 0
    RefreshFolders = True
End Function

Public Function RefreshLibrary() As Boolean
    Dim f As Form, sql As String, term As String, folderID As Long, state As String, selection As Long, filter As String
    Set f = Forms!frmLibrary
    selection = Nz(f!lstReports, 0)
    sql = "SELECT " & LibraryViewSelect() & " FROM Report WHERE 1=1"
    folderID = Nz(f!lstFolders, 0): state = Nz(f!cboArchive, "Active")
    If UXHasControl(f, "uxGroupOrganization") Then
        If folderID = -1 Then sql = sql & " AND ReportID NOT IN (SELECT ReportID FROM GroupMember)"
        If folderID > 0 Then sql = sql & " AND ReportID IN (SELECT ReportID FROM GroupMember WHERE GroupID=" & folderID & ")"
    Else
        If folderID = -1 Then sql = sql & " AND FolderID IS NULL"
        If folderID > 0 Then sql = sql & " AND FolderID=" & folderID
    End If
    If state = "Active" Then sql = sql & " AND IsArchived=False"
    If state = "Archived" Then sql = sql & " AND IsArchived=True"
    If UXHasControl(f, "uxTypeFilter") Then
        filter = Nz(f!uxTypeFilter, "All")
        If filter <> "All" Then sql = sql & " AND FormKind=" & QuoteSQL(filter)
        filter = Nz(f!uxStatusFilter, "All")
        If filter <> "All" Then sql = sql & " AND UXReportState(ReportID)=" & QuoteSQL(filter)
    End If
    term = Replace(Trim$(Nz(f!txtSearch, "")), "'", "''")
    If Len(term) > 0 Then sql = sql & " AND (InStr(1,Nz([Title],''),'" & term & "',1)>0 OR InStr(1,Nz([MemberName],''),'" & term & "',1)>0 OR InStr(1,Nz([OfficialName],''),'" & term & "',1)>0)"
    f!lstReports.RowSource = sql & " ORDER BY ModifiedAt DESC, ReportID DESC"
    f!lstReports.Requery
    If selection > 0 Then
        Dim i As Long
        For i = 1 To f!lstReports.ListCount - 1
            If Nz(f!lstReports.Column(0, i), 0) = selection Then f!lstReports = selection: Exit For
        Next
    End If
    Dim shown As Long
    shown = f!lstReports.ListCount - 1
    If shown < 0 Then shown = 0
    f!lblCount.Caption = shown & " reports shown"
    RefreshLibrary = True
End Function

Public Function AddFolderUI() As Boolean
    Dim name As String, parent As Long
    On Error GoTo Failed
    parent = Nz(Forms!frmLibrary!lstFolders, 0): If parent < 0 Then parent = 0
    name = InputBox("New folder name. A selected folder becomes its parent.", "Create organizational folder")
    If Len(Trim$(name)) = 0 Then Exit Function
    CreateFolder name, parent
    Call RefreshFolders()
    Call RefreshLibrary()
    AddFolderUI = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation
End Function

Public Function MoveReportUI() As Boolean
    Dim id As Long, destination As Long
    On Error GoTo Failed
    id = Nz(Forms!frmLibrary!lstReports, 0): destination = Nz(Forms!frmLibrary!cboDestination, 0)
    If id = 0 Then MsgBox "Select a report first.": Exit Function
    If CurrentProject.AllForms("frmEditor").IsLoaded Then MsgBox "Close the editor before moving a report.": Exit Function
    AssignFolder id, destination
    Call RefreshLibrary()
    MoveReportUI = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation
End Function

Public Function MoveFolderUI() As Boolean
    Dim id As Long, destination As Long
    On Error GoTo Failed
    id = Nz(Forms!frmLibrary!lstFolders, 0): destination = Nz(Forms!frmLibrary!cboDestination, 0)
    If id <= 0 Then MsgBox "Select an organizational folder first.": Exit Function
    MoveFolder id, destination
    Call RefreshFolders()
    Call RefreshLibrary()
    MoveFolderUI = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation
End Function

Public Function ShowHelp(ByVal topic As String) As Boolean
    DoCmd.OpenForm "frmHelp"
    Forms!frmHelp!lblTopic.Caption = topic
    Select Case topic
        Case "Folders"
            Forms!frmHelp!txtHelp = "Folders organize your local drafts. Select a folder to see its reports; Unassigned shows reports without a folder. Create folder uses the selected folder as its parent. Choose a destination, then Move report or Move folder. Moving organization does not change draft fields or official summary-group membership. All reports is a view, not a folder. Folder hierarchy uses native Access controls and requires no tree-view installation."
        Case "Saving"
            Forms!frmHelp!txtHelp = "Save incomplete drafts at any time. The Saved message confirms a committed save. Each changed save records contributor, time and previous/new values. A save with unchanged values adds no revision. Restore writes a new revision and preserves older history. Save before closing. Draft notes are not an official narrative block and their character count does not guarantee NAVFIT printed fit."
        Case Else
            Forms!frmHelp!txtHelp = "Swift is a local drafting workspace. Profiles label contributors; they are not authentication. Use one authoritative file, one person at a time. NAVFIT exchange and official validation remain unavailable until their compatibility proof passes. Final official printing remains in NAVFIT."
    End Select
    Forms!frmHelp.SetFocus: ShowHelp = True
End Function

Public Function CloseHelp() As Boolean
    DoCmd.Close acForm, "frmHelp", acSaveNo
    CloseHelp = True
End Function
