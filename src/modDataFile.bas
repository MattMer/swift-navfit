Attribute VB_Name = "modDataFile"
Option Compare Database
Option Explicit
' Where the workspace data is. Today the tables are in this file. When they are linked
' from a separate data file, that file is the workspace: its path, size, location and
' table-level reads are what matter. Code that needs the file itself asks here instead
' of assuming CurrentProject.
Private linkedData As DAO.Database, linkedDataPath As String

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
