Attribute VB_Name = "WorkspaceImportHarness"
Option Explicit
'Catch intentional failures inside VBA so Access does not show an unhandled
'runtime dialog when the external COM runner exercises rollback guards.
Public Function TryWorkspaceImport(ByVal path As String, Optional ByVal hash As String = "") As String
 On Error GoTo Failed
 If hash = "" Then hash = SourceFileHash(path)
 TryWorkspaceImport = ImportSwiftWorkspaceData(path, hash): Exit Function
Failed:
 TryWorkspaceImport = "REJECTED: " & Err.Number & " " & Err.Description
End Function
