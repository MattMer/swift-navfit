Attribute VB_Name = "modWorkspace"
Option Compare Database
Option Explicit

Public Function CreateFolder(ByVal folderName As String, Optional ByVal parentID As Long = 0) As Long
    Dim db As DAO.Database, r As DAO.Recordset
    Call ActiveContributorID
    folderName = Trim$(folderName)
    If Len(folderName) = 0 Or Len(folderName) > 80 Then Err.Raise vbObjectError + 200, "modWorkspace.CreateFolder", "Enter a folder name of 1 to 80 characters."
    If parentID <> 0 And DCount("*", "Folder", "FolderID=" & parentID) = 0 Then Err.Raise vbObjectError + 201, "modWorkspace.CreateFolder", "Parent folder does not exist."
    Set db = CurrentDb
    Set r = db.OpenRecordset("Folder", dbOpenDynaset)
    r.AddNew: r!FolderGuid = NewIdentity(): r!FolderName = folderName
    If parentID <> 0 Then r!ParentFolderID = parentID
    r.Update: r.Bookmark = r.LastModified: CreateFolder = r!FolderID: r.Close
End Function

Public Function FolderPath(ByVal folderID As Long) As String
    Dim parent As Long, depth As Long, part As String
    Do While folderID > 0
        If depth > 40 Then Err.Raise vbObjectError + 202, "modWorkspace.FolderPath", "Folder nesting exceeds supported depth."
        part = Nz(DLookup("FolderName", "Folder", "FolderID=" & folderID), "")
        If Len(part) = 0 Then Err.Raise vbObjectError + 203, "modWorkspace.FolderPath", "Folder path is incomplete."
        If Len(FolderPath) > 0 Then FolderPath = part & " / " & FolderPath Else FolderPath = part
        parent = Nz(DLookup("ParentFolderID", "Folder", "FolderID=" & folderID), 0)
        folderID = parent: depth = depth + 1
    Loop
End Function

'Presentation scope only. A refresh owns this map; writers keep FolderPath/live
'validation. Resolve lazily so an unrelated malformed folder is not newly read.
Public Function FolderPathSnapshot() As Object
    Dim result As Object, folders As Object, r As DAO.Recordset
    Set result = CreateObject("Scripting.Dictionary"): Set folders = CreateObject("Scripting.Dictionary")
    Set r = CurrentDb.OpenRecordset("SELECT FolderID,FolderName,ParentFolderID FROM Folder", dbOpenSnapshot)
    Do While Not r.EOF
        folders.Add CStr(r!FolderID), Array(r!FolderName.Value, Nz(r!ParentFolderID, 0)): r.MoveNext
    Loop
    r.Close: result.Add "Folders", folders: result.Add "Paths", CreateObject("Scripting.Dictionary")
    Set FolderPathSnapshot = result
End Function
Public Function FolderPathFromSnapshot(ByVal folderID As Long, ByVal snapshot As Object) As String
    Dim original As String, depth As Long, row As Variant, part As String
    original = CStr(folderID)
    If snapshot("Paths").Exists(original) Then FolderPathFromSnapshot = snapshot("Paths")(original): Exit Function
    Do While folderID > 0
        If depth > 40 Then Err.Raise vbObjectError + 202, "modWorkspace.FolderPathFromSnapshot", "Folder nesting exceeds supported depth."
        If Not snapshot("Folders").Exists(CStr(folderID)) Then Err.Raise vbObjectError + 203, "modWorkspace.FolderPathFromSnapshot", "Folder path is incomplete."
        row = snapshot("Folders")(CStr(folderID)): part = Nz(row(0), "")
        If Len(part) = 0 Then Err.Raise vbObjectError + 203, "modWorkspace.FolderPathFromSnapshot", "Folder path is incomplete."
        If Len(FolderPathFromSnapshot) > 0 Then FolderPathFromSnapshot = part & " / " & FolderPathFromSnapshot Else FolderPathFromSnapshot = part
        folderID = CLng(row(1)): depth = depth + 1
    Loop
    snapshot("Paths")(original) = FolderPathFromSnapshot
End Function

Public Function MoveFolder(ByVal folderID As Long, ByVal parentID As Long) As Boolean
    Dim cursor As Long, depth As Long, q As DAO.QueryDef
    Call ActiveContributorID
    If DCount("*", "Folder", "FolderID=" & folderID) = 0 Then Err.Raise vbObjectError + 204, "modWorkspace.MoveFolder", "Folder does not exist."
    cursor = parentID
    Do While cursor <> 0
        If cursor = folderID Then Err.Raise vbObjectError + 205, "modWorkspace.MoveFolder", "A folder cannot be moved into itself or its descendants."
        If DCount("*", "Folder", "FolderID=" & cursor) = 0 Then Err.Raise vbObjectError + 206, "modWorkspace.MoveFolder", "Destination folder does not exist."
        cursor = Nz(DLookup("ParentFolderID", "Folder", "FolderID=" & cursor), 0)
        depth = depth + 1
        If depth > 40 Then Err.Raise vbObjectError + 207, "modWorkspace.MoveFolder", "Folder path is too deep or contains a cycle."
    Loop
    Set q = CurrentDb.CreateQueryDef("", "PARAMETERS parent Long, id Long; UPDATE Folder SET ParentFolderID=parent WHERE FolderID=id")
    If parentID = 0 Then q!parent = Null Else q!parent = parentID
    q!id = folderID: q.Execute dbFailOnError: q.Close
    MoveFolder = True
End Function

Public Function AssignFolder(ByVal reportID As Long, ByVal folderID As Long, Optional ByVal auditOperationID As Long = 0) As Boolean
    Dim w As DAO.Workspace, db As DAO.Database, q As DAO.QueryDef, r As DAO.Recordset
    Dim oldFolder As Variant, pid As Long, failure As Long, message As String
    pid = ActiveContributorID()
    If folderID <> 0 And DCount("*", "Folder", "FolderID=" & folderID) = 0 Then Err.Raise vbObjectError + 209, "modWorkspace.AssignFolder", "Destination folder does not exist."
    Set w = DBEngine.Workspaces(0): Set db = CurrentDb
    On Error GoTo Failed
    w.BeginTrans
    Set r = db.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 210, "modWorkspace.AssignFolder", "Report does not exist."
    oldFolder = r!FolderID
    If r!IsArchived Then Err.Raise vbObjectError + 211, "modWorkspace.AssignFolder", "Recover the archived report before moving it."
    r.Close
    If Nz(oldFolder, 0) = folderID Then w.Rollback: AssignFolder = True: Exit Function
    Set q = db.CreateQueryDef("", "PARAMETERS folder Long, id Long; UPDATE Report SET FolderID=folder WHERE ReportID=id")
    If folderID = 0 Then q!folder = Null Else q!folder = folderID
    q!id = reportID: q.Execute dbFailOnError: q.Close
    Set r = db.OpenRecordset("FolderMove", dbOpenDynaset)
    r.AddNew: r!ReportID = reportID: r!OldFolderID = oldFolder
    If folderID <> 0 Then r!NewFolderID = folderID
    r!AuthorProfileID = pid: r!MovedAt = Now(): r.Update: r.Bookmark = r.LastModified
    AuditRecord "Folder move", CLng(r!FolderMoveID), auditOperationID: r.Close
    w.CommitTrans: AssignFolder = True: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modWorkspace", "AssignFolder", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description")): SwiftRollback w, swiftPrimary
    SwiftRaise swiftPrimary
End Function

Public Function RefreshFolders() As Boolean
    'Compatibility name; Reports is organized by summary group, not folder IDs.
    RefreshFolders = UXRefreshGroupFilter()
End Function

Public Function RefreshLibrary() As Boolean
    Dim f As Form, sql As String, term As String, folderID As Long, state As String, selection As Long, filter As String
    On Error GoTo Failed
    Set f = Forms!frmLibrary
    selection = Nz(f!lstReports, 0)
    sql = "SELECT " & LibraryViewSelect() & " FROM Report WHERE 1=1"
    folderID = Nz(f!lstFolders, 0): state = Nz(f!cboArchive, "Active")
    If folderID = -1 Then sql = sql & " AND ReportID NOT IN (SELECT ReportID FROM GroupMember)"
    If folderID > 0 Then sql = sql & " AND ReportID IN (SELECT ReportID FROM GroupMember WHERE GroupID=" & folderID & ")"
    If state = "Active" Then sql = sql & " AND IsArchived=False"
    If state = "Archived" Then sql = sql & " AND IsArchived=True"
    If UXHasControl(f, "uxTypeFilter") Then
        filter = Nz(f!uxTypeFilter, "All")
        If filter <> "All" Then sql = sql & " AND FormKind=" & QuoteSQL(filter)
        filter = Nz(f!uxStatusFilter, "All")
        If filter <> "All" Then sql = sql & " AND UXReportRowCell(ReportID,1)=" & QuoteSQL(filter)
        If UXHasControl(f, "cboRankFilter") Then sql = sql & UXMemberRankWhere(Nz(f!cboRankFilter, "All ranks/rates"))
    End If
    term = Left$(Trim$(Nz(f!txtSearch, "")), 255)
    If Len(term) > 0 Then sql = sql & " AND ReportSearchMatches(" & QuoteSQL(term) & ",[Title],[MemberName],[OfficialName])"
    sql = sql & ReportPeriodWhere(f!txtPeriodFrom, f!txtPeriodThrough)
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
    Exit Function
Failed:
    Forms!frmLibrary!lblCount.Caption = Err.Description
End Function

Public Function AddFolderUI() As Boolean
    'Orphan legacy callback: never interpret a summary-group ID as a folder ID.
    AddFolderUI = UXNavigate("frmBulkFolders")
End Function
Public Function MoveReportUI() As Boolean
    MoveReportUI = UXChooseReportGroup()
End Function
Public Function MoveFolderUI() As Boolean
    MoveFolderUI = UXNavigate("frmBulkFolders")
End Function

Public Function ShowHelp(ByVal topic As String) As Boolean
    Dim id As Long
    Select Case topic
        Case "Saving": id = 1
        Case "Folders": id = 2
        Case Else: id = 0
    End Select
    DoCmd.OpenForm "frmUXHelp"
    Forms!frmUXHelp!lstTopics = id
    ShowHelp = UXHelpSelected()
End Function

Public Function CloseHelp() As Boolean
    If CurrentProject.AllForms("frmUXHelp").IsLoaded Then DoCmd.Close acForm, "frmUXHelp", acSaveNo
    CloseHelp = True
End Function
