Attribute VB_Name = "ToolsHarness"
Option Explicit
Private checks As Long
Private Function Narrative(ByVal id As Long) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT OfficialComments FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Narrative = r!OfficialComments: r.Close
End Function
Private Function HasFinding(ByVal key As String) As Boolean
    Dim i As Long
    For i = 0 To Forms!frmChecklist!lstFindings.ListCount - 1
        If Forms!frmChecklist!lstFindings.Column(0, i) = key Then HasFinding = True
    Next
End Function
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 1500, , message
    checks = checks + 1
    CurrentDb.Execute "UPDATE AppMetadata SET CompatibilityStatus='TEST: " & Replace(message, "'", "''") & "' WHERE MetadataID=1", dbFailOnError
End Sub
Private Function Rejected(ByVal action As String, ByVal id As Long, Optional ByVal argument As String = "") As String
    On Error GoTo Failed
    Select Case action
        Case "snippet": Call InsertNarrativeSnippet(id)
        Case "undo": Call UndoNarrativeImport(id)
        Case "batch": Call UndoFolderBatch(id)
        Case "stale": Call AcceptIncomingNarrative(4, id, 1, argument)
        Case "source": Call AcceptIncomingNarrative(4, id, CLng(DLookup("CurrentRevisionID", "Report", "ReportID=" & id)), argument)
        Case "applybatch": Call ApplyFolderBatch(argument, id, FolderBatchStamp(argument))
    End Select
    Rejected = "NOT REJECTED": Exit Function
Failed:
    Rejected = Err.Description
End Function
Public Function TestTools(ByVal path As String) As String
    Dim id As Long, other As Long, revision As Long, snippet As Long, wrongSnippet As Long, receipt As Long
    Dim f As Form, before As Long, incoming As DAO.Database, r As DAO.Recordset
    Dim operationID As Long, folder As Long, selection As String, stamp As String, target As Long, hash As String, oldNarrative As Variant
    On Error GoTo Failed
    checks = 0: Call ConfigureWorkspace("SYNTHETIC TOOLS", "Synthetic Tester")
    id = CommitDraft(0, "EVAL", "", "", Null, Null, "Draft notes only", "Empty", Null)
    Check OpenDraftChecklist(id), "Checklist failed"
    Check Forms!frmChecklist!lstFindings.ListCount = 55, "Missing-input findings wrong"
    Check InStr(Forms!frmChecklist!lblSummary.Caption, "53 missing") > 0, "Missing count wrong"
    Forms!frmChecklist!lstFindings = "NARRATIVE": Check ChecklistGuidance(), "Guidance failed"
    Check InStr(Forms!frmChecklist!txtGuidance, "block 43") > 0, "Narrative guidance wrong"
    Call CloseDraftChecklist()
    Call CommitDraft(id, "EVAL", "Synthetic", "Synthetic Member", #10/2/2026#, #10/1/2026#, "Notes", "Dates", "Narrative")
    Check OpenDraftChecklist(id), "Checklist reload failed"
    Check InStr(Forms!frmChecklist!lblSummary.Caption, "1 draft date errors") > 0, "Reversed dates missed"
    Call CloseDraftChecklist()
    Call CommitDraft(id, "EVAL", "Synthetic", "Synthetic Member", #10/1/2026#, #10/2/2026#, "Notes", "Fix dates", "Narrative")
    revision = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
    oldNarrative = Narrative(id)
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!txtTitle = "Unsaved editor title": Call CheckpointEditor()
    Check ChecklistFromEditor(), "Editor checklist failed"
    Check InStr(Forms!frmChecklist!lblSummary.Caption, "48 missing") > 0, "Saved checklist used unsaved entries"
    Check Forms!frmChecklist!lstFindings.ListCount = 51, "Recovery finding missing"
    Call CloseDraftChecklist()
    Check Forms!frmEditor!txtTitle = "Unsaved editor title", "Checklist lost pending entry"
    Check DLookup("CurrentRevisionID", "Report", "ReportID=" & id) = revision, "Checklist changed history"
    snippet = SaveSnippet(0, "Synthetic phrase", "Snippet" & vbCrLf & String$(5000, "x"), "Synthetic command reference", "EVAL")
    wrongSnippet = SaveSnippet(0, "Officer phrase", "Wrong type", "Synthetic source", "FITREP")
    Check InStr(Rejected("snippet", wrongSnippet), "different report type") > 0, "Snippet type mismatch accepted"
    Check InsertNarrativeSnippet(snippet), "Snippet insert failed"
    Check InStr(Forms!frmEditor!txtOfficialComments, "Narrative" & vbCrLf & "Snippet") = 1, "Snippet did not append"
    Check DLookup("CurrentRevisionID", "Report", "ReportID=" & id) = revision, "Snippet committed automatically"
    Check SaveEditor(), "Snippet save failed"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    DoCmd.OpenForm "frmSnippets": Forms!frmSnippets!txtSearch = "phrase": Call LoadSnippets()
    Check Forms!frmSnippets!lstSnippets.ListCount = 2, "Snippet search failed"
    Forms!frmSnippets!txtSearch = "'": Call LoadSnippets()
    Check Forms!frmSnippets!lstSnippets.ListCount = 0, "Snippet apostrophe search failed"
    DoCmd.Close acForm, "frmSnippets", acSaveNo
    DoCmd.OpenForm "frmLibrary"
    Check SaveLibraryView("Dates"), "Library view save failed": Call RefreshLibrary()
    Check InStr(Forms!frmLibrary!lstReports.RowSource, "PeriodStart") > 0, "Date columns missing"
    Check SaveLibraryView("History"), "History view save failed": Call RefreshLibrary()
    Check InStr(Forms!frmLibrary!lstReports.RowSource, "CurrentRevisionID") > 0, "History column missing"
    Check SaveLibraryView("Drafting"), "Drafting view restore failed": Call RefreshLibrary()
    Call OpenReviewFile(path): hash = SourceFileHash(path)
    Check PreviewIncomingNarrative(4, id), "Reconcile preview failed"
    Set incoming = DBEngine.OpenDatabase(path, False, True)
    Set r = incoming.OpenRecordset("SELECT Comments FROM Reports WHERE ReportID=4", dbOpenSnapshot)
    Check IsNull(Forms!frmReconcile!txtIncoming) = IsNull(r!Comments), "Incoming preview null wrong"
    revision = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
    Check InStr(Rejected("source", id, "WRONG HASH"), "Source changed") > 0, "Changed-source preview accepted"
    Check InStr(Rejected("stale", id, hash), "Destination changed") > 0, "Stale destination accepted"
    target = AcceptIncomingNarrative(4, id, revision, hash)
    Check target = id, "Reconcile created another report"
    Check IsNull(Narrative(id)) = IsNull(r!Comments), "Accepted narrative null differs"
    If Not IsNull(r!Comments) Then Check StrComp(Narrative(id), r!Comments, vbBinaryCompare) = 0, "Accepted narrative differs from source"
    Check DLookup("Title", "Report", "ReportID=" & id) = "Unsaved editor title", "Reconcile changed other fields"
    receipt = DLookup("ImportReceiptID", "ImportReceipt", "ReportID=" & id)
    Check DLookup("BeforeRevisionID", "ImportReceipt", "ImportReceiptID=" & receipt) = revision, "Reconcile baseline not retained"
    Check UndoNarrativeImport(receipt), "Reconcile undo failed"
    'The baseline for reconciliation includes the inserted snippet, not the earlier draft.
    Check StrComp(Nz(Narrative(id), ""), "Narrative" & vbCrLf & "Snippet" & vbCrLf & String$(5000, "x"), vbBinaryCompare) = 0, "Undo did not restore exact prior narrative"
    Check DLookup("CurrentRevisionID", "Report", "ReportID=" & id) <> revision, "Undo erased rather than appended history"
    Check DLookup("IsUndone", "ImportReceipt", "ImportReceiptID=" & receipt), "Undo receipt not marked"
    Check InStr(Rejected("undo", receipt), "already undone") > 0, "Repeat undo accepted"
    Call CloseReconcileUI(): r.Close: incoming.Close
    other = ImportReviewedNarrative(2)
    receipt = DLookup("ImportReceiptID", "ImportReceipt", "ReportID=" & other)
    Check UndoNarrativeImport(receipt), "New import undo failed"
    Check DLookup("IsArchived", "Report", "ReportID=" & other), "New import was not archived"
    other = ImportReviewedNarrative(3)
    receipt = DLookup("ImportReceiptID", "ImportReceipt", "ReportID=" & other)
    Call CommitDraft(other, "CHIEFEVAL", "Later title", "", Null, Null, "")
    Check InStr(Rejected("undo", receipt), "Later edits") > 0, "Later edits overwritten by undo"
    Check OpenDraftChecklist(other), "Imported checklist failed"
    Check HasFinding("PARTIAL_IMPORT"), "Import completeness missing"
    Call CloseDraftChecklist()
    folder = CreateFolder("Batch destination")
    selection = CStr(id) & "," & CStr(other): stamp = FolderBatchStamp(selection)
    before = DCount("*", "FolderMove")
    CurrentDb.TableDefs("FolderBatchChange").Fields("ReportID").ValidationRule = "<> " & other
    Check Rejected("applybatch", folder, selection) <> "NOT REJECTED", "Injected batch failure accepted"
    Check DCount("*", "FolderMove") = before And IsNull(DLookup("FolderID", "Report", "ReportID=" & id)) And IsNull(DLookup("FolderID", "Report", "ReportID=" & other)), "Batch failure left partial moves"
    CurrentDb.TableDefs("FolderBatchChange").Fields("ReportID").ValidationRule = ""
    operationID = ApplyFolderBatch(selection, folder, stamp)
    Check DCount("*", "FolderBatchChange", "OperationID=" & operationID) = 2, "Batch entries missing"
    Check DLookup("FolderID", "Report", "ReportID=" & id) = folder, "Batch move failed"
    Check UndoFolderBatch(operationID), "Batch undo failed"
    Check IsNull(DLookup("FolderID", "Report", "ReportID=" & id)), "Batch undo lost unassigned state"
    operationID = ApplyFolderBatch(selection, folder, FolderBatchStamp(selection))
    Call AssignFolder(id, 0)
    Check InStr(Rejected("batch", operationID), "Later changes") > 0, "Later move did not block entire undo"
    Check DLookup("FolderID", "Report", "ReportID=" & other) = folder, "Rejected undo partially changed batch"
    Check DCount("*", "OperationJournal", "OperationState='Failed'") > 0, "Failure journal missing"
    Check DCount("*", "OperationJournal", "OperationState='Completed'") > 0, "Completion journal missing"
    operationID = BeginOperation("Synthetic interrupted", "Test only")
    DoCmd.OpenForm "frmOperations": Forms!frmOperations!lstOperations = operationID: Call OperationDetails()
    Check InStr(Forms!frmOperations!txtOutcome, "Started or interrupted") > 0, "Incomplete entry not explained"
    TestTools = "PASS: " & checks & " integrated tools assertions": Exit Function
Failed:
    TestTools = "FAIL: " & Err.Number & " " & Err.Description
End Function
