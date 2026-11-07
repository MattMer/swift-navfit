Attribute VB_Name = "modOperations"
Option Compare Database
Option Explicit
Private lastFailure As Object

'Keep originals in the owned in-memory context, never in a diagnostic log.
'Capture before any cleanup; numeric error codes remain backward compatible.
Public Function SwiftFailure(ByVal module As String, ByVal procedure As String, ByVal stage As String, ByVal number As Long, ByVal description As String, ByVal source As String, Optional ByVal transaction As String = "Not started") As Object
    Dim result As Object
    If Len(stage) = 0 Then stage = "Validate / initialize"
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "Identity", module & "." & procedure & ":" & number
    result.Add "Module", module: result.Add "Procedure", procedure
    result.Add "Stage", stage: result.Add "Number", number
    result.Add "Description", description: result.Add "Source", source
    result.Add "Transaction", transaction: result.Add "Cleanup", ""
    Set lastFailure = result: Set SwiftFailure = result
End Function

Public Function SwiftRedact(ByVal text As String) As String
    Dim i As Long, last As Long, digits As Long, result As String
    'Identifiers may be compact, hyphenated or space-separated. Deliberately
    'redact long digit sequences too; logs do not need full personal identifiers.
    'A run is 9 or more digits, each optionally preceded by one space or hyphen.
    'Plain VBA: Microsoft is retiring VBScript, which provided RegExp.
    i = 1
    Do While i <= Len(text)
        If IsDigitAt(text, i) Then
            last = DigitRunEnd(text, i, digits)
            If digits >= 9 Then result = result & "[identifier redacted]" Else result = result & Mid$(text, i, last - i + 1)
            i = last + 1
        Else
            result = result & Mid$(text, i, 1): i = i + 1
        End If
    Loop
    SwiftRedact = result
End Function
' Longest run of consecutive digits in text (no separators).
Public Function LongestDigitRun(ByVal text As String) As Long
    Dim i As Long, run As Long
    For i = 1 To Len(text)
        If IsDigitAt(text, i) Then run = run + 1 Else run = 0
        If run > LongestDigitRun Then LongestDigitRun = run
    Next
End Function
Public Function IsDigitAt(ByVal text As String, ByVal i As Long) As Boolean
    Dim c As String
    If i < 1 Or i > Len(text) Then Exit Function
    c = Mid$(text, i, 1): IsDigitAt = (c >= "0" And c <= "9")
End Function
' From a digit at start, the end of a run in which each further digit may follow
' one space or hyphen; digits returns how many digits the run holds.
Private Function DigitRunEnd(ByVal text As String, ByVal start As Long, ByRef digits As Long) As Long
    Dim k As Long, c As String
    digits = 1: k = start
    Do
        If IsDigitAt(text, k + 1) Then
            k = k + 1: digits = digits + 1
        Else
            c = Mid$(text, k + 1, 1)
            If (c = " " Or c = "-") And IsDigitAt(text, k + 2) Then k = k + 2: digits = digits + 1 Else Exit Do
        End If
    Loop
    DigitRunEnd = k
End Function

Public Function SwiftFailureText(ByVal context As Object) As String
    Set lastFailure = context
    SwiftFailureText = SwiftRedact(context("Description")) & vbCrLf & "Diagnostic: " & context("Identity") & "; source: " & SwiftRedact(context("Source")) & "; stage: " & SwiftRedact(context("Stage")) & "; transaction: " & context("Transaction")
    If Len(context("Cleanup")) > 0 Then SwiftFailureText = SwiftFailureText & vbCrLf & "Cleanup: " & context("Cleanup")
End Function

Public Function SwiftLastFailure() As String
    If Not lastFailure Is Nothing Then SwiftLastFailure = SwiftFailureText(lastFailure)
End Function
Public Function SwiftLastFailureContext() As Object
    'Return scalar diagnostic facts without exposing the owned mutable context.
    Dim result As Object, key As Variant
    If lastFailure Is Nothing Then Exit Function
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In lastFailure.Keys: result.Add CStr(key), lastFailure(key): Next
    Set SwiftLastFailureContext = result
End Function

Public Sub SwiftCleanupError(ByVal context As Object, ByVal step As String, ByVal number As Long)
    If Len(context("Cleanup")) > 0 Then context("Cleanup") = context("Cleanup") & "; "
    context("Cleanup") = context("Cleanup") & step & " (" & number & ")"
End Sub

Public Sub SwiftRollback(ByVal workspace As DAO.workspace, ByVal context As Object)
    On Error GoTo Failed
    workspace.Rollback
    context("Transaction") = "Rolled back"
    Exit Sub
Failed:
    context("Transaction") = "Rollback failed; inspect saved state before retry"
    SwiftCleanupError context, "rollback", Err.number
End Sub

Public Sub SwiftClose(ByVal owned As Object, ByVal context As Object, ByVal role As String)
    On Error GoTo Failed
    If Not owned Is Nothing Then owned.Close
    Exit Sub
Failed:
    SwiftCleanupError context, role, Err.number
End Sub

Public Sub SwiftFinishFailed(ByVal operationID As Long, ByVal context As Object)
    On Error GoTo Failed
    If operationID > 0 Then FinishOperation operationID, "Failed", SwiftFailureText(context)
    Exit Sub
Failed:
    SwiftCleanupError context, "record failed operation", Err.number
End Sub

Public Sub SwiftRaise(ByVal context As Object)
    Set lastFailure = context
    Err.Raise CLng(context("Number")), CStr(context("Module")) & "." & context("Procedure"), SwiftFailureText(context)
End Sub
Public Function BeginOperation(ByVal kind As String, ByVal source As String) As Long
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset, pid As Long, w As DAO.Workspace, failure As Long, message As String
    pid = ActiveContributorID()
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = ownedDB.OpenRecordset("OperationJournal", dbOpenDynaset)
    r.AddNew: r!OperationGuid = NewIdentity(): r!OperationKind = kind: r!AuthorProfileID = pid
    r!SourceDescription = source: r!StartedAt = Now(): r!OperationState = "Started"
    r.Update: r.Bookmark = r.LastModified: BeginOperation = r!OperationID: r.Close
    AuditRecord "Operation", BeginOperation
    w.CommitTrans: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modOperations", "BeginOperation", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary: On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
'Journal entries move forward only: Started may finish as Completed, Failed or
'Reviewed, or pause as Prepared (files written, not yet published); Prepared may
'then finish. Completed, Failed and Reviewed are final.
Public Function OperationMayBecome(ByVal current As String, ByVal target As String) As Boolean
    Select Case current
        Case "Started": OperationMayBecome = (target = "Completed" Or target = "Failed" Or target = "Reviewed" Or target = "Prepared")
        Case "Prepared": OperationMayBecome = (target = "Completed" Or target = "Failed" Or target = "Reviewed")
    End Select
End Function
Public Sub FinishOperation(ByVal operationID As Long, ByVal state As String, ByVal message As String)
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset, w As DAO.Workspace, failure As Long, explanation As String
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = ownedDB.OpenRecordset("SELECT * FROM OperationJournal WHERE OperationID=" & operationID, dbOpenDynaset)
    If r.EOF Then Err.Raise vbObjectError + 1101, "modOperations.FinishOperation", "Operation entry missing."
    If Not OperationMayBecome(Nz(r!OperationState, ""), state) Then Err.Raise vbObjectError + 1108, "modOperations.FinishOperation", "Operation " & operationID & " is already " & Nz(r!OperationState, "unknown") & "; it cannot become " & state & "."
    r.Edit: r!OperationState = state
    If state = "Failed" Then r!Outcome = SwiftRedact(message) Else r!Outcome = message
    r!FinishedAt = Now(): r.Update: r.Close
    AuditRecord "Operation", operationID, operationID, False, True
    w.CommitTrans: Exit Sub
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modOperations", "FinishOperation", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): explanation = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary: On Error GoTo 0
    SwiftRaise swiftPrimary
End Sub
Public Function UndoNarrativeImport(ByVal receiptID As Long) As Boolean
    Dim r As DAO.Recordset, w As DAO.Workspace, id As Long, operationID As Long, trans As Boolean
    Dim failure As Long, message As String, stage As String, beforeRevision As Variant
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ImportReceipt WHERE ImportReceiptID=" & receiptID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 1102, "modOperations.UndoNarrativeImport", "Import receipt does not exist."
    operationID = BeginOperation("Undo import", "Receipt " & receiptID)
    id = r!ReportID
    If Nz(r!IsUndone, False) Then Err.Raise vbObjectError + 1103, "modOperations.UndoNarrativeImport", "Import was already undone."
    If IsNull(r!AppliedRevisionID) Then Err.Raise vbObjectError + 1104, "modOperations.UndoNarrativeImport", "This older receipt lacks an applied revision. Automatic undo is unavailable."
    If Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & id), 0) <> r!AppliedRevisionID Then Err.Raise vbObjectError + 1105, "modOperations.UndoNarrativeImport", "Later edits exist. Undo stopped to preserve them."
    If Nz(DLookup("IsArchived", "Report", "ReportID=" & id), False) Or Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") <> "Draft" Then Err.Raise vbObjectError + 1106, "modOperations.UndoNarrativeImport", "Report is archived or locked. Undo stopped."
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 1107, "modOperations.UndoNarrativeImport", "Close the editor before undoing an import."
    beforeRevision = r!BeforeRevisionID: r.Close
    stage = "Begin undo transaction"
    Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    If IsNull(beforeRevision) Then
        stage = "Archive imported draft"
        Call ArchiveDraft(id, True, operationID)
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
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modOperations", "UndoNarrativeImport", stage, Err.Number, Err.Description, Err.Source, IIf(trans, "Active", "Not active"))
    failure = CLng(swiftPrimary("Number")): message = stage & ": " & CStr(swiftPrimary("Description"))
    If trans Then SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    If operationID > 0 Then SwiftFinishFailed operationID, swiftPrimary
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
Public Function LoadOperations() As Boolean
    Forms!frmOperations!lstOperations.RowSource = "SELECT OperationID,OperationKind,OperationState,StartedAt FROM OperationJournal ORDER BY OperationID DESC"
    Forms!frmOperations!lstOperations.Requery
    Forms!frmOperations!lstReceipts.RowSource = "SELECT ImportReceiptID,ReportID,ExternalReportID,ImportedAt,IsUndone FROM ImportReceipt ORDER BY ImportReceiptID DESC"
    Forms!frmOperations!lstReceipts.Requery
    Forms!frmOperations!lstAudit.RowSource = "SELECT AuditID,EventAt AS [When],EventKind AS [Event],ActorCaption AS [Contributor],ReportID AS [Report],CorrelationGuid AS [Operation key] FROM AuditEvent ORDER BY AuditID DESC"
    Forms!frmOperations!lstAudit.Requery
    LoadOperations = True
End Function
Public Function OperationDetails() As Boolean
    Dim r As DAO.Recordset, id As Long
    id = Nz(Forms!frmOperations!lstOperations, 0): If id = 0 Then Exit Function
    Set r = CurrentDb.OpenRecordset("SELECT * FROM OperationJournal WHERE OperationID=" & id, dbOpenSnapshot)
    If Not r.EOF Then Forms!frmOperations!txtOutcome = r!OperationState & vbCrLf & Nz(r!Outcome, "Started or interrupted. Inspect saved reports and receipts before retrying.") & vbCrLf & Nz(r!SourceDescription, "")
    r.Close: OperationDetails = True
End Function
Public Function AcknowledgeOperationUI() As Boolean
    Dim id As Long
    id = Nz(Forms!frmOperations!lstOperations, 0)
    Select Case Nz(DLookup("OperationState", "OperationJournal", "OperationID=" & id), "")
        Case "Started"
            If MsgBox("Acknowledge this incomplete entry after reviewing reports and import receipts? This does not retry or undo it.", vbOKCancel + vbQuestion) <> vbOK Then Exit Function
        Case "Prepared"
            If MsgBox("Acknowledge this interrupted export after checking its folder for unpublished files? This does not retry or remove them.", vbOKCancel + vbQuestion) <> vbOK Then Exit Function
        Case Else
            Exit Function
    End Select
    Call FinishOperation(id, "Reviewed", "Incomplete operation acknowledged by active contributor. No data changes performed.")
    Call LoadOperations(): AcknowledgeOperationUI = True
End Function
