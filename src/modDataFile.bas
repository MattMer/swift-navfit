Attribute VB_Name = "modDataFile"
Option Compare Database
Option Explicit
' Where the workspace data is. Today the tables are in this file. When they are linked
' from a separate data file, that file is the workspace: its path, size, location and
' table-level reads are what matter. Code that needs the file itself asks here instead
' of assuming CurrentProject.
Private linkedData As DAO.Database, linkedDataPath As String, otherSessionAtOpen As Boolean

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
' A session that holds the file makes the next one's exclusive open fail, which is the
' test. Returns True when another session had it open.
Public Function SwiftDataHold() As Boolean
    Dim probe As DAO.Database, filePath As String, failure As Long, message As String
    If Not SwiftDataIsLinked() Then Exit Function
    filePath = SwiftDataPath()
    If Not linkedData Is Nothing And StrComp(linkedDataPath, filePath, vbTextCompare) = 0 Then SwiftDataHold = otherSessionAtOpen: Exit Function
    Set linkedData = Nothing: otherSessionAtOpen = False
    On Error Resume Next
    Set probe = DBEngine.Workspaces(0).OpenDatabase(filePath, True)
    failure = Err.Number: message = Err.Description
    On Error GoTo 0
    Select Case failure
        Case 0
            probe.Close
        Case 3045, 3356, 3704
            'In use, or already opened by a user on some machine: someone else has it.
            otherSessionAtOpen = True
        Case Else
            Err.Raise failure, "modDataFile.SwiftDataHold", message
    End Select
    Set linkedData = DBEngine.Workspaces(0).OpenDatabase(filePath): linkedDataPath = filePath
    SwiftDataHold = otherSessionAtOpen
End Function

' Lets go of this session's hold on the data file, for an upgrade that needs it alone.
Public Sub SwiftDataRelease()
    If Not linkedData Is Nothing Then linkedData.Close
    Set linkedData = Nothing: linkedDataPath = ""
End Sub

' True when another session had the data file open as this one started.
Public Function SwiftDataOtherSession() As Boolean
    SwiftDataOtherSession = otherSessionAtOpen
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
Public Function SwiftLockConflict(ByVal number As Long) As Boolean
    Select Case number
        Case 3008, 3009, 3046, 3186, 3187, 3188, 3197, 3202, 3211, 3212, 3218, 3260, 3262, 3624
            SwiftLockConflict = True
    End Select
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
