Attribute VB_Name = "modBulkFolders"
Option Compare Database
Option Explicit
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
    Dim id As Variant, r As DAO.Recordset
    For Each id In IDs(selection)
        Set r = CurrentDb.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
        If r.EOF Then Err.Raise vbObjectError + 1402, , "Selected report missing."
        If r!IsArchived Or r!WorkflowState <> "Draft" Then Err.Raise vbObjectError + 1403, , "Selected report is archived or locked."
        FolderBatchStamp = FolderBatchStamp & id & ":" & Nz(r!FolderID, 0) & ":" & Nz(r!CurrentRevisionID, 0) & ":" & Nz(DMax("FolderMoveID", "FolderMove", "ReportID=" & id), 0) & ";"
        r.Close
    Next
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
    Dim failure As Long, message As String
    On Error GoTo Failed
    undoID = BeginOperation("Undo folder batch", "Operation " & operationID)
    Set r = CurrentDb.OpenRecordset("SELECT * FROM FolderBatchChange WHERE OperationID=" & operationID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 1406, , "No folder changes for this operation."
    Do While Not r.EOF
        id = r!ReportID
        If r!IsUndone Then Err.Raise vbObjectError + 1407, , "Batch already undone."
        If Nz(DMax("FolderMoveID", "FolderMove", "ReportID=" & id), 0) <> r!AppliedMoveID Or Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & id), 0) <> r!BaseRevisionID Then Err.Raise vbObjectError + 1408, , "Later changes exist. Entire batch undo stopped."
        If Nz(DLookup("IsArchived", "Report", "ReportID=" & id), False) Or Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") <> "Draft" Then Err.Raise vbObjectError + 1409, , "Archived or locked report blocks batch undo."
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
    Forms!frmBulkFolders!lstReports.RowSource = "SELECT ReportID,Title,UXIdentity(OfficialName,MemberName) AS Member,IIf(FolderID Is Null,'Unassigned',FolderPath(Nz(FolderID,0))) AS CurrentFolder FROM Report WHERE IsArchived=False AND WorkflowState='Draft' ORDER BY Title"
    Forms!frmBulkFolders!cboDestination.RowSource = "SELECT 0 AS ID,'Unassigned' AS Label FROM AppMetadata UNION ALL SELECT FolderID,FolderPath(FolderID) FROM Folder"
    Forms!frmBulkFolders!txtStamp = "": LoadBulkFolders = True
End Function
Private Function SelectedReports() As String
    Dim item As Variant
    For Each item In Forms!frmBulkFolders!lstReports.ItemsSelected
        If Len(SelectedReports) > 0 Then SelectedReports = SelectedReports & ","
        SelectedReports = SelectedReports & Forms!frmBulkFolders!lstReports.ItemData(item)
    Next
End Function
Public Function PreviewFolderBatchUI() As Boolean
    Dim selection As String, id As Variant, text As String
    On Error GoTo Failed
    selection = SelectedReports()
    Forms!frmBulkFolders!txtStamp = FolderBatchStamp(selection)
    Forms!frmBulkFolders!txtSelection = selection: Forms!frmBulkFolders!txtDestination = Nz(Forms!frmBulkFolders!cboDestination, 0)
    text = "Destination: " & IIf(Nz(Forms!frmBulkFolders!cboDestination, 0) = 0, "Unassigned", FolderPath(Nz(Forms!frmBulkFolders!cboDestination, 0))) & vbCrLf
    For Each id In IDs(selection)
        text = text & "Report " & id & ": " & Nz(DLookup("Title", "Report", "ReportID=" & id), "") & " | " & IIf(Nz(DLookup("FolderID", "Report", "ReportID=" & id), 0) = 0, "Unassigned", FolderPath(Nz(DLookup("FolderID", "Report", "ReportID=" & id), 0))) & vbCrLf
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
