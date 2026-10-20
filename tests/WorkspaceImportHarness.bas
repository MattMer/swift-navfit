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

Public Function ProbeImportedCounter() As Boolean
 Dim db As DAO.Database, rows As DAO.Recordset, w As DAO.Workspace, before As Long, inserted As Long, trans As Boolean
 On Error GoTo Failed
 Set db = CurrentDb: Set w = DBEngine.Workspaces(0)
 before = DMax("RevisionValueID", "RevisionValue")
 w.BeginTrans: trans = True
 Set rows = db.OpenRecordset("RevisionValue", dbOpenDynaset)
 rows.AddNew: rows!RevisionID = DMin("RevisionID", "ReportRevision")
 rows!FieldKey = "BulkCounterProbe": rows!ValueText = "Synthetic probe": rows!IsNullValue = False
 rows.Update: rows.Bookmark = rows.LastModified: inserted = rows!RevisionValueID: rows.Close
 w.Rollback: trans = False
 ProbeImportedCounter = (inserted > before And DMax("RevisionValueID", "RevisionValue") = before)
 Exit Function
Failed:
 On Error Resume Next
 If Not rows Is Nothing Then rows.Close
 If trans Then w.Rollback
End Function
