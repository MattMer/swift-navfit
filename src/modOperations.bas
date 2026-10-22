Attribute VB_Name = "modOperations"
Option Compare Database
Option Explicit
Public Function BeginOperation(ByVal kind As String, ByVal source As String) As Long
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset, pid As Long
    pid = ActiveContributorID()
    Set r = ownedDB.OpenRecordset("OperationJournal", dbOpenDynaset)
    r.AddNew: r!OperationGuid = NewIdentity(): r!OperationKind = kind: r!AuthorProfileID = pid
    r!SourceDescription = source: r!StartedAt = Now(): r!OperationState = "Started"
    r.Update: r.Bookmark = r.LastModified: BeginOperation = r!OperationID: r.Close
End Function
Public Sub FinishOperation(ByVal operationID As Long, ByVal state As String, ByVal message As String)
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset
    Set r = ownedDB.OpenRecordset("SELECT * FROM OperationJournal WHERE OperationID=" & operationID, dbOpenDynaset)
    If r.EOF Then Err.Raise vbObjectError + 1101, , "Operation entry missing."
    r.Edit: r!OperationState = state: r!Outcome = message: r!FinishedAt = Now(): r.Update: r.Close
End Sub
Public Function UndoNarrativeImport(ByVal receiptID As Long) As Boolean
    Dim r As DAO.Recordset, w As DAO.Workspace, id As Long, operationID As Long, trans As Boolean
    Dim failure As Long, message As String, stage As String, beforeRevision As Variant
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ImportReceipt WHERE ImportReceiptID=" & receiptID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 1102, , "Import receipt does not exist."
    operationID = BeginOperation("Undo import", "Receipt " & receiptID)
    id = r!ReportID
    If Nz(r!IsUndone, False) Then Err.Raise vbObjectError + 1103, , "Import was already undone."
    If IsNull(r!AppliedRevisionID) Then Err.Raise vbObjectError + 1104, , "This older receipt lacks an applied revision. Automatic undo is unavailable."
    If Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & id), 0) <> r!AppliedRevisionID Then Err.Raise vbObjectError + 1105, , "Later edits exist. Undo stopped to preserve them."
    If Nz(DLookup("IsArchived", "Report", "ReportID=" & id), False) Or Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") <> "Draft" Then Err.Raise vbObjectError + 1106, , "Report is archived or locked. Undo stopped."
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 1107, , "Close the editor before undoing an import."
    beforeRevision = r!BeforeRevisionID: r.Close
    stage = "Begin undo transaction"
    Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    If IsNull(beforeRevision) Then
        stage = "Archive imported draft"
        Call ArchiveDraft(id, True)
    Else
        stage = "Restore preceding revision"
        Call RestoreDraft(id, CLng(beforeRevision))
    End If
    stage = "Mark receipt undone"
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ImportReceipt WHERE ImportReceiptID=" & receiptID, dbOpenDynaset)
    r.Edit: r!IsUndone = True: r.Update: r.Close
    Call FinishOperation(operationID, "Completed", "Compensated receipt " & receiptID & ". History and provenance retained.")
    w.CommitTrans: trans = False: UndoNarrativeImport = True: Exit Function
Failed:
    failure = Err.Number: message = stage & ": " & Err.Description
    On Error Resume Next
    If trans Then w.Rollback
    If Not r Is Nothing Then r.Close
    If operationID > 0 Then FinishOperation operationID, "Failed", message
    On Error GoTo 0
    Err.Raise failure, "UndoNarrativeImport", message
End Function
Public Function LoadOperations() As Boolean
    Forms!frmOperations!lstOperations.RowSource = "SELECT OperationID,OperationKind,OperationState,StartedAt FROM OperationJournal ORDER BY OperationID DESC"
    Forms!frmOperations!lstOperations.Requery
    Forms!frmOperations!lstReceipts.RowSource = "SELECT ImportReceiptID,ReportID,ExternalReportID,ImportedAt,IsUndone FROM ImportReceipt ORDER BY ImportReceiptID DESC"
    Forms!frmOperations!lstReceipts.Requery
    LoadOperations = True
End Function
Public Function OperationDetails() As Boolean
    Dim r As DAO.Recordset, id As Long
    id = Nz(Forms!frmOperations!lstOperations, 0): If id = 0 Then Exit Function
    Set r = CurrentDb.OpenRecordset("SELECT * FROM OperationJournal WHERE OperationID=" & id, dbOpenSnapshot)
    If Not r.EOF Then Forms!frmOperations!txtOutcome = r!OperationState & vbCrLf & Nz(r!Outcome, "Started or interrupted. Inspect saved reports and receipts before retrying.") & vbCrLf & Nz(r!SourceDescription, "")
    r.Close: OperationDetails = True
End Function
Public Function UndoImportUI() As Boolean
    On Error GoTo Failed
    If MsgBox("Undo the selected import? New imported drafts are archived. Reconciled narratives are restored through a new revision. Later edits block undo; provenance stays retained.", vbOKCancel + vbQuestion) <> vbOK Then Exit Function
    UndoImportUI = UndoNarrativeImport(Nz(Forms!frmOperations!lstReceipts, 0))
    Call LoadOperations()
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Undo import"
End Function
Public Function AcknowledgeOperationUI() As Boolean
    Dim id As Long
    id = Nz(Forms!frmOperations!lstOperations, 0)
    If Nz(DLookup("OperationState", "OperationJournal", "OperationID=" & id), "") <> "Started" Then Exit Function
    If MsgBox("Acknowledge this incomplete entry after reviewing reports and import receipts? This does not retry or undo it.", vbOKCancel + vbQuestion) <> vbOK Then Exit Function
    Call FinishOperation(id, "Reviewed", "Incomplete operation acknowledged by active contributor. No data changes performed.")
    Call LoadOperations(): AcknowledgeOperationUI = True
End Function
