Attribute VB_Name = "BuildCleanupHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise 5, , message
 checks = checks + 1
End Sub
Public Function TestBuildCleanup() As String
 Dim folder As Long, child As Long, id As Long, revision As Long, before As Long, moves As Long, other As Long, operation As Long, selection As String, stamp As String, failure As Long
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC BUILD CLEANUP", "Build tester"
 folder = CreateFolder("Preserved organizational folder")
 child = CreateFolder("Child folder", folder)
 id = CommitDraft(0, "EVAL", "Cleanup fixture", "SYNTHETIC MEMBER", Null, Null, "")
 revision = CurrentRevision(id)
 Check AssignFolder(id, folder), "Folder assignment service lost"
 moves = DCount("*", "FolderMove")
 Check MoveFolder(child, 0), "Folder hierarchy service lost"
 Check FolderPath(child) = "Child folder", "Folder path wrong"
 other = CommitDraft(0, "EVAL", "Batch fixture", "SYNTHETIC SECOND", Null, Null, "")
 selection = CStr(id) & "," & CStr(other): stamp = FolderBatchStamp(selection)
 before = DCount("*", "FolderMove")
 CurrentDb.TableDefs("FolderBatchChange").Fields("ReportID").ValidationRule = "<> " & other
 On Error Resume Next
 ApplyFolderBatch selection, child, stamp
 failure = Err.Number: Err.Clear: On Error GoTo Failed
 CurrentDb.TableDefs("FolderBatchChange").Fields("ReportID").ValidationRule = ""
 Check failure <> 0, "Injected folder batch failure was accepted"
 Check DCount("*", "FolderMove") = before And DLookup("FolderID", "Report", "ReportID=" & id) = folder And IsNull(DLookup("FolderID", "Report", "ReportID=" & other)), "Failed folder batch left partial moves"
 Check DCount("*", "OperationJournal", "OperationState='Failed'") > 0, "Failed batch journal missing"
 operation = ApplyFolderBatch(selection, child, stamp)
 Check DCount("*", "FolderBatchChange", "OperationID=" & operation) = 2, "Folder batch history missing"
 Check DLookup("FolderID", "Report", "ReportID=" & id) = child And DLookup("FolderID", "Report", "ReportID=" & other) = child, "Folder batch did not move both reports"
 Check UndoFolderBatch(operation), "Folder batch undo failed"
 Check DLookup("FolderID", "Report", "ReportID=" & id) = folder And IsNull(DLookup("FolderID", "Report", "ReportID=" & other)), "Undo lost original folder assignments"
 operation = ApplyFolderBatch(selection, child, FolderBatchStamp(selection))
 AssignFolder id, folder
 On Error Resume Next
 UndoFolderBatch operation
 failure = Err.Number: Err.Clear: On Error GoTo Failed
 Check failure <> 0 And DLookup("FolderID", "Report", "ReportID=" & other) = child, "Stale undo partially changed folder assignments"
 operation = BeginOperation("Synthetic interrupted", "Fixture only")
 DoCmd.OpenForm "frmOperations": Forms!frmOperations!lstOperations = operation: OperationDetails
 Check InStr(Forms!frmOperations!txtOutcome, "Started or interrupted") > 0, "Interrupted operation not explained"
 DoCmd.Close acForm, "frmOperations", acSaveNo
 TestBuildCleanup = "PASS: " & checks & " build cleanup and folder-history checks": Exit Function
Failed:
 TestBuildCleanup = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
