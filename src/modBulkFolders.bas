Attribute VB_Name = "modBulkFolders"
Option Compare Database
Option Explicit
Private displayFolders As Object
Private Function IDs(ByVal text As String) As Collection
    Dim list As New Collection, part As Variant, id As Long
    If Len(text) = 0 Then Err.Raise vbObjectError + 1400, , "Select reports first."
    For Each part In Split(text, ",")
        id = CLng(part)
        If id <= 0 Or CStr(id) <> CStr(part) Then Err.Raise vbObjectError + 1401, , "Invalid report selection."
        list.Add id, CStr(id)
    Next
    Set IDs = list
End Function
Public Function FolderBatchStamp(ByVal selection As String) As String
    Dim selected As Collection, rows As Object
    Set selected = IDs(selection): Set rows = FolderReadRows(selected)
    FolderBatchStamp = FolderRowsStamp(selected, rows)
End Function
Private Function FolderRowsStamp(ByVal selected As Collection, ByVal rows As Object) As String
    Dim id As Variant, row As Variant
    For Each id In selected
        row = rows(CStr(id))
        If Not row(0) Then Err.Raise vbObjectError + 1402, , "Selected report missing."
        If row(4) Or row(5) <> "Draft" Then Err.Raise vbObjectError + 1403, , "Selected report is archived or locked."
        FolderRowsStamp = FolderRowsStamp & id & ":" & Nz(row(2), 0) & ":" & Nz(row(3), 0) & ":" & row(6) & ";"
    Next
End Function
Private Function FolderReadRows(ByVal selected As Collection) As Object
    Dim result As Object, db As DAO.Database, r As DAO.Recordset, i As Long, last As Long, csv As String, j As Long, row As Variant
    Set result = CreateObject("Scripting.Dictionary"): Set db = CurrentDb
    For i = 1 To selected.Count Step 200
        csv = "": last = i + 199: If last > selected.Count Then last = selected.Count
        For j = i To last
            If csv <> "" Then csv = csv & ","
            csv = csv & CLng(selected(j))
            result.Add CStr(selected(j)), Array(False, Null, Null, Null, Null, Null, 0)
        Next
        Set r = db.OpenRecordset("SELECT ReportID,Title,FolderID,CurrentRevisionID,IsArchived,WorkflowState FROM Report WHERE ReportID In (" & csv & ")", dbOpenSnapshot)
        Do While Not r.EOF
            result(CStr(r!ReportID)) = Array(True, r!Title.Value, r!FolderID.Value, r!CurrentRevisionID.Value, r!IsArchived.Value, r!WorkflowState.Value, 0): r.MoveNext
        Loop
        r.Close
        Set r = db.OpenRecordset("SELECT ReportID,Max(FolderMoveID) AS LastMove FROM FolderMove WHERE ReportID In (" & csv & ") GROUP BY ReportID", dbOpenSnapshot)
        Do While Not r.EOF
            row = result(CStr(r!ReportID)): row(6) = Nz(r!LastMove, 0): result(CStr(r!ReportID)) = row: r.MoveNext
        Loop
        r.Close
    Next
    Set FolderReadRows = result
End Function
Public Function ApplyFolderBatch(ByVal selection As String, ByVal destination As Long, ByVal expected As String) As Long
    Dim operationID As Long, id As Variant, before As Variant, afterMove As Long
    Dim r As DAO.Recordset, moved As DAO.Recordset, w As DAO.Workspace, trans As Boolean, failure As Long, message As String
    On Error GoTo Failed
    operationID = BeginOperation("Bulk folder move", selection & " to folder " & destination)
    If destination <> 0 And DCount("*", "Folder", "FolderID=" & destination) = 0 Then Err.Raise vbObjectError + 1404, , "Destination missing."
    If expected <> FolderBatchStamp(selection) Then Err.Raise vbObjectError + 1405, , "Reports changed after preview. Preview again."
    Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    For Each id In IDs(selection)
        before = DLookup("FolderID", "Report", "ReportID=" & id)
        If Nz(before, 0) <> destination Then
            Call AssignFolder(CLng(id), destination, operationID)
            Set moved = CurrentDb.OpenRecordset("SELECT Max(FolderMoveID) AS Latest FROM FolderMove WHERE ReportID=" & id, dbOpenSnapshot)
            afterMove = moved!Latest: moved.Close
            Set r = CurrentDb.OpenRecordset("FolderBatchChange", dbOpenDynaset)
            r.AddNew: r!OperationID = operationID: r!ReportID = id: r!OldFolderID = before
            If destination <> 0 Then r!NewFolderID = destination
            r!AppliedMoveID = afterMove: r!BaseRevisionID = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
            r!IsUndone = False: r.Update: r.Close
        End If
    Next
    Call FinishOperation(operationID, "Completed", "Folder moves committed. Official report revisions unchanged.")
    w.CommitTrans: trans = False: ApplyFolderBatch = operationID: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If trans Then w.Rollback
    If Not r Is Nothing Then r.Close
    If operationID > 0 Then FinishOperation operationID, "Failed", message
    On Error GoTo 0
    Err.Raise failure, "ApplyFolderBatch", message
End Function
Public Function UndoFolderBatch(ByVal operationID As Long) As Boolean
    Dim r As DAO.Recordset, w As DAO.Workspace, undoID As Long, trans As Boolean, id As Long
    Dim failure As Long, message As String, selected As New Collection, seen As Object, rows As Object, row As Variant
    On Error GoTo Failed
    undoID = BeginOperation("Undo folder batch", "Operation " & operationID)
    Set r = CurrentDb.OpenRecordset("SELECT * FROM FolderBatchChange WHERE OperationID=" & operationID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 1406, , "No folder changes for this operation."
    Set seen = CreateObject("Scripting.Dictionary")
    Do While Not r.EOF
        If Not seen.Exists(CStr(r!ReportID)) Then seen.Add CStr(r!ReportID), True: selected.Add CLng(r!ReportID)
        r.MoveNext
    Loop
    Set rows = FolderReadRows(selected): r.MoveFirst
    Do While Not r.EOF
        id = r!ReportID
        If r!IsUndone Then Err.Raise vbObjectError + 1407, , "Batch already undone."
        row = rows(CStr(id))
        If row(6) <> r!AppliedMoveID Or Nz(row(3), 0) <> r!BaseRevisionID Then Err.Raise vbObjectError + 1408, , "Later changes exist. Entire batch undo stopped."
        If Nz(row(4), False) Or Nz(row(5), "") <> "Draft" Then Err.Raise vbObjectError + 1409, , "Archived or locked report blocks batch undo."
        r.MoveNext
    Loop
    r.Close
    Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    Set r = CurrentDb.OpenRecordset("SELECT * FROM FolderBatchChange WHERE OperationID=" & operationID, dbOpenDynaset)
    Do While Not r.EOF
        Call AssignFolder(CLng(r!ReportID), CLng(Nz(r!OldFolderID, 0)), undoID)
        r.Edit: r!IsUndone = True: r.Update: r.MoveNext
    Loop
    r.Close: Call FinishOperation(undoID, "Completed", "Compensating folder moves recorded for operation " & operationID)
    w.CommitTrans: trans = False: UndoFolderBatch = True: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If trans Then w.Rollback
    If Not r Is Nothing Then r.Close
    If undoID > 0 Then FinishOperation undoID, "Failed", message
    On Error GoTo 0
    Err.Raise failure, "UndoFolderBatch", message
End Function
Public Function LoadBulkFolders() As Boolean
    Set displayFolders = FolderPathSnapshot()
    Forms!frmBulkFolders!lstReports.RowSource = "SELECT ReportID,Title,UXIdentity(OfficialName,MemberName) AS Member,IIf(FolderID Is Null,'Unassigned',BulkFolderPath(Nz(FolderID,0))) AS CurrentFolder FROM Report WHERE IsArchived=False AND WorkflowState='Draft' ORDER BY Title"
    Forms!frmBulkFolders!cboDestination.RowSource = "SELECT 0 AS ID,'Unassigned' AS Label FROM AppMetadata UNION ALL SELECT FolderID,BulkFolderPath(FolderID) FROM Folder"
    Forms!frmBulkFolders!txtStamp = "": LoadBulkFolders = True
End Function
Public Function BulkFolderPath(ByVal folderID As Long) As String
    If displayFolders Is Nothing Then Set displayFolders = FolderPathSnapshot()
    BulkFolderPath = FolderPathFromSnapshot(folderID, displayFolders)
End Function
Private Function SelectedReports() As String
    Dim item As Variant
    For Each item In Forms!frmBulkFolders!lstReports.ItemsSelected
        If Len(SelectedReports) > 0 Then SelectedReports = SelectedReports & ","
        SelectedReports = SelectedReports & Forms!frmBulkFolders!lstReports.ItemData(item)
    Next
End Function
Public Function PreviewFolderBatchUI() As Boolean
    Dim selection As String, id As Variant, text As String, selected As Collection, rows As Object, folders As Object, row As Variant
    On Error GoTo Failed
    selection = SelectedReports()
    Set selected = IDs(selection): Set rows = FolderReadRows(selected): Set folders = FolderPathSnapshot()
    Forms!frmBulkFolders!txtStamp = FolderRowsStamp(selected, rows)
    Forms!frmBulkFolders!txtSelection = selection: Forms!frmBulkFolders!txtDestination = Nz(Forms!frmBulkFolders!cboDestination, 0)
    text = "Destination: " & IIf(Nz(Forms!frmBulkFolders!cboDestination, 0) = 0, "Unassigned", FolderPathFromSnapshot(Nz(Forms!frmBulkFolders!cboDestination, 0), folders)) & vbCrLf
    For Each id In selected
        row = rows(CStr(id))
        text = text & "Report " & id & ": " & Nz(row(1), "") & " | " & IIf(Nz(row(2), 0) = 0, "Unassigned", FolderPathFromSnapshot(Nz(row(2), 0), folders)) & vbCrLf
    Next
    Forms!frmBulkFolders!txtPreview = text & vbCrLf & "Only organizational folder assignments change. Official revisions and groups stay unchanged."
    PreviewFolderBatchUI = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Preview folder moves"
End Function
Public Function ApplyFolderBatchUI() As Boolean
    On Error GoTo Failed
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 1410, , "Close the editor first."
    If Nz(Forms!frmBulkFolders!txtSelection, "") <> SelectedReports() Or Nz(Forms!frmBulkFolders!txtDestination, -1) <> Nz(Forms!frmBulkFolders!cboDestination, 0) Then Err.Raise vbObjectError + 1411, , "Selection changed. Preview again."
    If MsgBox("Apply the previewed folder moves as one batch?", vbOKCancel + vbQuestion) <> vbOK Then Exit Function
    Call ApplyFolderBatch(SelectedReports(), CLng(Nz(Forms!frmBulkFolders!cboDestination, 0)), Nz(Forms!frmBulkFolders!txtStamp, ""))
    Call LoadBulkFolders(): Forms!frmBulkFolders!txtPreview = "Batch completed. Operations screen provides guarded undo."
    If CurrentProject.AllForms("frmLibrary").IsLoaded Then Call RefreshLibrary()
    ApplyFolderBatchUI = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Apply folder moves"
End Function
Public Function UndoFolderBatchUI() As Boolean
    On Error GoTo Failed
    If MsgBox("Undo all folder moves in this operation? Any later report or folder move blocks the entire undo.", vbOKCancel + vbQuestion) <> vbOK Then Exit Function
    UndoFolderBatchUI = UndoFolderBatch(Nz(Forms!frmOperations!lstOperations, 0)): Call LoadOperations(): Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Undo folder batch"
End Function
