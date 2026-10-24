Attribute VB_Name = "modBulkFolders"
Option Compare Database
Option Explicit
Private displayFolders As Object
Private Function IDs(ByVal text As String) As Collection
    Dim list As New Collection, part As Variant, id As Long
    If Len(text) = 0 Then Err.Raise vbObjectError + 1400, "modBulkFolders.IDs", "Select reports first."
    For Each part In Split(text, ",")
        id = CLng(part)
        If id <= 0 Or CStr(id) <> CStr(part) Then Err.Raise vbObjectError + 1401, "modBulkFolders.IDs", "Invalid report selection."
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
        If Not row(0) Then Err.Raise vbObjectError + 1402, "modBulkFolders.FolderRowsStamp", "Selected report missing."
        If row(4) Or row(5) <> "Draft" Then Err.Raise vbObjectError + 1403, "modBulkFolders.FolderRowsStamp", "Selected report is archived or locked."
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
    If destination <> 0 And DCount("*", "Folder", "FolderID=" & destination) = 0 Then Err.Raise vbObjectError + 1404, "modBulkFolders.ApplyFolderBatch", "Destination missing."
    If expected <> FolderBatchStamp(selection) Then Err.Raise vbObjectError + 1405, "modBulkFolders.ApplyFolderBatch", "Reports changed after preview. Preview again."
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
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modBulkFolders", "ApplyFolderBatch", "Execute", Err.Number, Err.Description, Err.Source, IIf(trans, "Active", "Not active"))
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    If trans Then SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    If operationID > 0 Then SwiftFinishFailed operationID, swiftPrimary
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
Public Function UndoFolderBatch(ByVal operationID As Long) As Boolean
    Dim r As DAO.Recordset, w As DAO.Workspace, undoID As Long, trans As Boolean, id As Long
    Dim failure As Long, message As String, selected As New Collection, seen As Object, rows As Object, row As Variant
    On Error GoTo Failed
    undoID = BeginOperation("Undo folder batch", "Operation " & operationID)
    Set r = CurrentDb.OpenRecordset("SELECT * FROM FolderBatchChange WHERE OperationID=" & operationID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 1406, "modBulkFolders.UndoFolderBatch", "No folder changes for this operation."
    Set seen = CreateObject("Scripting.Dictionary")
    Do While Not r.EOF
        If Not seen.Exists(CStr(r!ReportID)) Then seen.Add CStr(r!ReportID), True: selected.Add CLng(r!ReportID)
        r.MoveNext
    Loop
    Set rows = FolderReadRows(selected): r.MoveFirst
    Do While Not r.EOF
        id = r!ReportID
        If r!IsUndone Then Err.Raise vbObjectError + 1407, "modBulkFolders.UndoFolderBatch", "Batch already undone."
        row = rows(CStr(id))
        If row(6) <> r!AppliedMoveID Or Nz(row(3), 0) <> r!BaseRevisionID Then Err.Raise vbObjectError + 1408, "modBulkFolders.UndoFolderBatch", "Later changes exist. Entire batch undo stopped."
        If Nz(row(4), False) Or Nz(row(5), "") <> "Draft" Then Err.Raise vbObjectError + 1409, "modBulkFolders.UndoFolderBatch", "Archived or locked report blocks batch undo."
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
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modBulkFolders", "UndoFolderBatch", "Execute", Err.Number, Err.Description, Err.Source, IIf(trans, "Active", "Not active"))
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    If trans Then SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    If undoID > 0 Then SwiftFinishFailed undoID, swiftPrimary
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
Public Function LoadBulkFolders() As Boolean
    'Retired UI callbacks preserve safe navigation without changing folder data.
    LoadBulkFolders = UXNavigate("frmLibrary")
End Function
Public Function BulkFolderPath(ByVal folderID As Long) As String
    If displayFolders Is Nothing Then Set displayFolders = FolderPathSnapshot()
    BulkFolderPath = FolderPathFromSnapshot(folderID, displayFolders)
End Function
Public Function PreviewFolderBatchUI() As Boolean
    PreviewFolderBatchUI = UXNavigate("frmLibrary")
End Function
Public Function ApplyFolderBatchUI() As Boolean
    ApplyFolderBatchUI = UXNavigate("frmLibrary")
End Function
Public Function UndoFolderBatchUI() As Boolean
    On Error GoTo Failed
    If MsgBox("Undo all folder moves in this operation? Any later report or folder move blocks the entire undo.", vbOKCancel + vbQuestion) <> vbOK Then Exit Function
    UndoFolderBatchUI = UndoFolderBatch(Nz(Forms!frmOperations!lstOperations, 0)): Call LoadOperations(): Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Undo folder batch"
End Function
