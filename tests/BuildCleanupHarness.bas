Attribute VB_Name = "BuildCleanupHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise 5, , message
 checks = checks + 1
End Sub
Public Function TestBuildCleanup() As String
 Dim folder As Long, child As Long, id As Long, revision As Long, before As Long, moves As Long, group As Long, f As Access.Form, hasLegacy As Boolean, item As AccessObject, other As Long, operation As Long, selection As String, stamp As String, failure As Long
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
 group = CreateSummaryGroup("Summary group with overlapping ID", "EVAL", "E5")
 DoCmd.OpenForm "frmLibrary": Forms!frmLibrary!lstReports = id
 Forms!frmLibrary!lstFolders = group
 before = DCount("*", "Folder")
 Check AddFolderUI(), "Legacy add-folder adapter failed"
 Check Screen.ActiveForm.Name = "frmLibrary", "Retired folder callback did not return to Reports"
 Check Not UXHasControl(Forms!frmLibrary, "cmdBulkFolders"), "Reports still offers retired Folder tools"
 Check PreviewFolderBatchUI() And ApplyFolderBatchUI() And Screen.ActiveForm.Name = "frmLibrary", "Retired folder actions did not return safely to Reports"
 Check DCount("*", "Folder") = before And DCount("*", "FolderMove") = moves, "Group ID was misinterpreted as a folder mutation"
 Check CurrentRevision(id) = revision, "Folder adapter changed report revision"
 Check OpenScreen("frmLibrary"), "Return to Reports failed"
 Check Screen.ActiveForm.Name = "frmLibrary", "Reports did not receive focus"
 Check MoveFolderUI(), "Legacy move-folder adapter failed"
 Check DCount("*", "FolderMove") = moves, "Orphan callback moved a folder"
 Check OpenScreen("frmHelp"), "Legacy help alias failed"
 Check Screen.ActiveForm.Name = "frmUXHelp", "Legacy help opened wrong form"
 CloseHelp
 Check OpenScreen("frmLibrary"), "Reports return failed"
 Forms!frmLibrary!lstReports = id
 Check MoveReportUI(), "Legacy report-move callback failed"
 Check CurrentProject.AllForms("frmUXMoveGroup").IsLoaded, "Report move did not use group preview"
 DoCmd.Close acForm, "frmUXMoveGroup", acSaveNo
 For Each item In CurrentProject.AllForms
  If item.Name = "frmHelp" Or item.Name = "frmUXReview" Or item.Name = "frmReviewComments" Or item.Name = "frmChecklist" Or item.Name = "frmUXNoteComposer" Or item.Name = "frmBulkFolders" Then hasLegacy = True
 Next
 Check Not hasLegacy, "Retired standalone form still installed"
 DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
 Forms!frmEditor!txtOfficialName = "SAVED THROUGH LEGACY NAV"
 Check OpenScreen("frmLibrary"), "Legacy main-menu entry did not leave report safely"
 Check Not CurrentProject.AllForms("frmEditor").IsLoaded, "Legacy route left old editor open"
 Check DLookup("OfficialName", "Report", "ReportID=" & id) = "SAVED THROUGH LEGACY NAV", "Legacy route lost pending report values"
 Check DCount("*", "FolderMove") = moves, "Navigation changed folder history"
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
 TestBuildCleanup = "PASS: " & checks & " build cleanup, folder-history and guarded legacy-route checks": Exit Function
Failed:
 TestBuildCleanup = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
