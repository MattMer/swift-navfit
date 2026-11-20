Attribute VB_Name = "modDataFile"
Option Compare Database
Option Explicit
' Where the workspace data is. Today the tables are in this file. When they are linked
' from a separate data file, that file is the workspace: its path, size, location and
' table-level reads are what matter. Code that needs the file itself asks here instead
' of assuming CurrentProject.
Private linkedData As DAO.Database, linkedDataPath As String
'An ADO connection opened with record-level locking before anything else opens the data
'file. The first program to open an Access file sets its locking mode for everyone on it,
'and a plain DAO open sets page-level locking, where people saving different reports
'block each other. Held for the session; released for an upgrade.
Private rowLockHold As Object, rowLockPath As String, rowLockError As String
' Tries for a save whose records another session holds (see SwiftLockWait).
Public Const SWIFT_LOCK_TRIES As Long = 8

' Full path of the file that holds the workspace tables.
Public Function SwiftDataPath() As String
    Dim source As String, position As Long
    source = CurrentDb.TableDefs("Workspace").Connect
    position = InStr(1, source, "DATABASE=", vbTextCompare)
    If position = 0 Then SwiftDataPath = CurrentProject.FullName: Exit Function
    source = Mid$(source, position + 9)
    position = InStr(source, ";")
    If position > 0 Then source = Left$(source, position - 1)
    SwiftDataPath = source
End Function

' True when the workspace tables are linked from another file.
Public Function SwiftDataIsLinked() As Boolean
    SwiftDataIsLinked = StrComp(SwiftDataPath(), CurrentProject.FullName, vbTextCompare) <> 0
End Function

' True for this application file and for the file that holds its data: neither can be
' an import source, an export destination or a staging file.
Public Function SwiftIsWorkspaceFile(ByVal filePath As String) As Boolean
    SwiftIsWorkspaceFile = StrComp(filePath, CurrentProject.FullName, vbTextCompare) = 0 Or StrComp(filePath, SwiftDataPath(), vbTextCompare) = 0
End Function

' With a separate data file, called once at startup: finds out whether another session
' has the data file open, then keeps it open for this session. Access opens a linked
' file only while it reads or writes, so without this no session could see another.
' Keep a shared handle on the data file for the session. It never opens the file
' exclusively, so people starting Swift NAVFIT at the same moment cannot collide.
' Returns True when another session is connected now.
Public Function SwiftDataHold() As Boolean
    Dim filePath As String
    If Not SwiftDataIsLinked() Then Exit Function
    filePath = SwiftDataPath()
    If linkedData Is Nothing Or StrComp(linkedDataPath, filePath, vbTextCompare) <> 0 Then
        Set linkedData = Nothing
        Set linkedData = SwiftOpenDataFile(filePath, False): linkedDataPath = filePath
    End If
    SwiftRowLockHold filePath
    SwiftLockPatience
    SwiftDataHold = SwiftDataOtherSession()
End Function

' Open the data file, waiting a few seconds while another session briefly has it busy
' (opening, upgrading or compacting) instead of reporting it as missing.
Public Function SwiftOpenDataFile(ByVal filePath As String, ByVal readOnly As Boolean) As DAO.Database
    Dim attempt As Long, failure As Long, message As String
    'Only this session's own data file is held; a file being checked is left alone.
    If StrComp(filePath, SwiftDataPath(), vbTextCompare) = 0 Then SwiftRowLockHold filePath
    For attempt = 1 To 6
        On Error Resume Next
        Set SwiftOpenDataFile = DBEngine.Workspaces(0).OpenDatabase(filePath, False, readOnly)
        failure = Err.Number: message = Err.Description
        On Error GoTo 0
        If failure = 0 Then Exit Function
        Select Case failure
            Case 3045, 3050, 3051, 3356, 3704, 3734, 3006, 3008, 3009, 3260
                If attempt < 6 Then SwiftPause 500
            Case Else
                Exit For
        End Select
    Next
    Err.Raise failure, "modDataFile.SwiftOpenDataFile", message
End Function

' Open the data file in record-level locking mode, once per session, so this session
' never makes it page-locked for everyone. Best effort: a file that is missing, busy or
' read-only is left for the caller's own open to report.
Public Sub SwiftRowLockHold(ByVal filePath As String)
    Dim c As Object
    If Not rowLockHold Is Nothing And StrComp(rowLockPath, filePath, vbTextCompare) = 0 Then Exit Sub
    On Error GoTo Done
    If Not rowLockHold Is Nothing Then rowLockHold.Close
    Set rowLockHold = Nothing: rowLockPath = ""
    Set c = CreateObject("ADODB.Connection")
    c.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Jet OLEDB:Database Locking Mode=1"
    Set rowLockHold = c: rowLockPath = filePath: rowLockError = ""
    Exit Sub
Done:
    rowLockError = Err.Number & " " & Err.Description
End Sub

' The record-level hold's state, for tests: "" when none, else its last error or "held".
Public Function SwiftRowLockHeld() As String
    If rowLockHold Is Nothing Then SwiftRowLockHeld = "none: " & rowLockError Else SwiftRowLockHeld = "held " & rowLockPath
End Function

' Row-locked writes for tables every save touches. An UPDATE or DELETE statement locks
' whole pages, so people saving different reports would hold each other up; editing or
' deleting through a recordset locks only the records it changes.
Public Sub SwiftDeleteRows(ByVal db As DAO.Database, ByVal table As String, ByVal where As String)
    Dim r As DAO.Recordset
    Set r = db.OpenRecordset("SELECT * FROM [" & table & "] WHERE " & where, dbOpenDynaset)
    Do Until r.EOF
        r.Delete
        r.MoveNext
    Loop
    r.Close
End Sub
Public Sub SwiftSetField(ByVal db As DAO.Database, ByVal table As String, ByVal field As String, ByVal value As Variant, ByVal where As String)
    Dim r As DAO.Recordset
    Set r = db.OpenRecordset("SELECT * FROM [" & table & "] WHERE " & where, dbOpenDynaset)
    Do Until r.EOF
        r.Edit: r.Fields(field).Value = value: r.Update
        r.MoveNext
    Loop
    r.Close
End Sub

' Saves take turns adding revision rows, so a save made while someone else's is
' finishing waits in place for up to about 12 seconds (120 engine retries) instead of
' failing after 2. This session only: the setting is not written to the registry.
Public Sub SwiftLockPatience()
    On Error Resume Next
    DBEngine.SetOption 57, 120
End Sub

' Lets go of this session's hold on the data file, for an upgrade that needs it alone.
Public Sub SwiftDataRelease()
    On Error Resume Next
    If Not linkedData Is Nothing Then linkedData.Close
    Set linkedData = Nothing: linkedDataPath = ""
    If Not rowLockHold Is Nothing Then rowLockHold.Close
    Set rowLockHold = Nothing: rowLockPath = ""
End Sub

' True when another session is connected to this workspace now.
Public Function SwiftDataOtherSession() As Boolean
    Dim names As String
    SwiftDataOtherSession = (PresenceOthersNow(names) > 0)
End Function

' The database that owns the tables, for table-type recordsets and Seek, which linked
' tables refuse. With local tables it is the caller's database, unchanged. With linked
' tables it is the data file opened in the default workspace, so it shares the caller's
' transaction.
Public Function SwiftDataDb(Optional ByVal db As DAO.Database) As DAO.Database
    Dim filePath As String
    filePath = SwiftDataPath()
    If StrComp(filePath, CurrentProject.FullName, vbTextCompare) = 0 Then
        If db Is Nothing Then Set SwiftDataDb = CurrentDb Else Set SwiftDataDb = db
        Exit Function
    End If
    If linkedData Is Nothing Or StrComp(linkedDataPath, filePath, vbTextCompare) <> 0 Then
        Set linkedData = DBEngine.Workspaces(0).OpenDatabase(filePath)
        linkedDataPath = filePath
    End If
    Set SwiftDataDb = linkedData
End Function

' Workspace tables: not Access's own (MSys, ~) and not the blank set a linked copy
' keeps to make a new data file from (USysBlank, see modDataLink).
Public Function SwiftWorkspaceTable(ByVal tableName As String) As Boolean
    SwiftWorkspaceTable = Left$(tableName, 4) <> "MSys" And Left$(tableName, 1) <> "~" And Left$(tableName, 4) <> "USys"
End Function

' Errors Access raises when another session holds a lock on the same records or tables.
' Read what other sessions have saved, not this session's cached copy, before a check
' that depends on it.
Public Sub SwiftFreshRead()
    On Error Resume Next
    DBEngine.Idle dbRefreshCache
End Sub
' How many times a save is tried when another session holds its records, and the wait
' before each retry: growing, with a random part so sessions do not retry in step.
Public Sub SwiftLockWait(ByVal attempt As Long)
    Randomize
    SwiftPause 300 * attempt + CLng(Rnd() * 500)
End Sub
Public Function SwiftLockConflict(ByVal number As Long) As Boolean
    Select Case number
        Case 3008, 3009, 3046, 3167, 3186, 3187, 3188, 3197, 3202, 3211, 3212, 3218, 3260, 3262, 3624
            SwiftLockConflict = True
    End Select
End Function

' The text to show for the current error: a save another session was holding up gets
' the plain explanation, anything else its own description.
Public Function SwiftErrorText() As String
    If SwiftLockConflict(Err.Number) Then SwiftErrorText = SwiftLockMessage() Else SwiftErrorText = Err.Description
End Function
Public Function SwiftLockMessage() As String
    SwiftLockMessage = "Someone else was saving to the same records at that moment, so this change was not saved. Nothing was lost: your entries are still here. Try again in a moment."
End Function

' A short wait that keeps the screen responsive.
Public Sub SwiftPause(ByVal milliseconds As Long)
    Dim started As Single
    started = Timer
    Do While (Timer - started) * 1000 < milliseconds And Timer >= started
        DoEvents
    Loop
End Sub

' Tables that only describe this moment (who has the file open). They are linked like
' any other, but an import never copies them and a backup does not check them.
Public Function SwiftTransientTable(ByVal tableName As String) As Boolean
    SwiftTransientTable = (tableName = "SessionPresence")
End Function
