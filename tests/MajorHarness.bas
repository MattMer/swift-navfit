Attribute VB_Name = "MajorHarness"
Option Compare Binary
Option Explicit
Private count As Long
Private Sub Check(ByVal value As Boolean, ByVal message As String)
    If Not value Then Err.Raise vbObjectError + 4900, , message
    count = count + 1
End Sub
Public Function TestMajor() As String
    Dim a As Long, b As Long, g As Long, snap As Long, meeting As Long, proposal As Long, batch As Long, route As Long, template As Long, rev As Long, failed As Boolean, undo As Long, suggestion As Object, secondProfile As Long, firstProfile As Long, siblingRoute As Long
    On Error GoTo Failed
    count = 0
    Call ConfigureWorkspace("SYNTHETIC COMMAND", "Synthetic contributor")
    a = CommitDraft(0, "EVAL", "Synthetic A", "Example A", Null, Null, "", "Test")
    b = CommitDraft(0, "EVAL", "Synthetic B", "Example B", Null, Null, "", "Test")
    batch = PreviewFieldBatch(a & "," & b, "OfficialComments", "O'Brien & synthetic", "Reviewed batch")
    Check Nz(MajorScalar("SELECT OfficialComments FROM Report WHERE ReportID=" & a), "") = "", "Preview does not commit values"
    ApplyFieldBatch batch
    Check MajorScalar("SELECT OfficialComments FROM Report WHERE ReportID=" & a) = "O'Brien & synthetic", "Batch writes first report"
    Check MajorScalar("SELECT OfficialComments FROM Report WHERE ReportID=" & b) = "O'Brien & synthetic", "Batch writes second report"
    Check MajorScalar("SELECT BatchState FROM ChangeBatch WHERE BatchID=" & batch) = "Applied", "Batch records completion"
    batch = PreviewFieldBatch(a & "," & b, "Title", "Reviewed title", "Stale check")
    Call CommitDraft(b, "EVAL", "Concurrent edit", "Example B", Null, Null, "", "Concurrent")
    rev = CurrentRevision(a)
    On Error Resume Next
    ApplyFieldBatch batch
    failed = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    Check failed, "Stale second report rejects entire batch"
    Check CurrentRevision(a) = rev, "Rollback restores first report revision"
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & a) = "Synthetic A", "Rollback restores first report value"
    Check MajorScalar("SELECT BatchState FROM ChangeBatch WHERE BatchID=" & batch) = "Preview", "Failed batch remains pending"
    g = CreateSummaryGroup("Synthetic E5 group", "EVAL", "E5")
    AddGroupMember g, a, "E5"
    AddGroupMember g, b, "E5"
    snap = FreezeSummaryGroup(g, "Synthetic roster checked against test list")
    Check DCount("*", "GroupSnapshotMember", "SnapshotID=" & snap) = 2, "Frozen roster contains both saved reports"
    Check MajorScalar("SELECT Stamp FROM GroupSnapshot WHERE SnapshotID=" & snap) = GroupStamp(g), "Frozen stamp matches current group"
    meeting = CreateComparisonSet(snap, "Synthetic meeting")
    proposal = ProposeMeetingField(meeting, a, "Title", "Meeting decision")
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & a) = "Synthetic A", "Proposal does not edit report"
    ApplyMeetingProposal proposal
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & a) = "Meeting decision", "Explicit proposal apply creates edit"
    Check MajorScalar("SELECT Stamp FROM GroupSnapshot WHERE SnapshotID=" & snap) <> GroupStamp(g), "Edit makes group snapshot stale"
    Check MajorScalar("SELECT RevisionID FROM ComparisonMember WHERE SetID=" & meeting & " AND ReportID=" & a) = rev, "Meeting remains bound to original revision"
    template = CreateRouteTemplate("Synthetic one-step route", CStr(MajorAuthor()), "Check saved report")
    route = StartRoute(a, template, g)
    Check RouteIsCurrent(route), "New route is current"
    DecideRouteStep route, "Accepted", "Synthetic review complete"
    Check MajorScalar("SELECT RunState FROM RouteRun WHERE RunID=" & route) = "Accepted", "Last accepted step completes route"
    Check RouteIsCurrent(route), "Accepted route remains current before edits"
    batch = PreviewFieldBatch(CStr(a), "Title", "After approval", "Material edit")
    ApplyFieldBatch batch
    Check Not RouteIsCurrent(route), "Material edit makes approval stale"
    Check MajorScalar("SELECT RunState FROM RouteRun WHERE RunID=" & route) = "Accepted", "Historical approval preserved"
    Check DCount("*", "GroupSnapshotMember", "SnapshotID=" & snap) = 2, "Frozen roster unchanged after edits"
    batch = PreviewFieldBatch(a & "," & b, "Title", "After approval", "First no-op, then change")
    rev = CurrentRevision(a): ApplyFieldBatch batch
    Check CurrentRevision(a) = rev, "No-op item creates no history"
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & b) = "After approval", "No-op item does not roll back later batch work"
    undo = PreviewBatchUndo(batch)
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & b) = "After approval", "Undo preview does not replace values"
    ApplyFieldBatch undo
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & b) = "Concurrent edit", "Undo restores original value as new revision"
    SetFieldDefault "Command", 1, "ShipStation", "COMMAND DEFAULT"
    Set suggestion = ResolveFieldDefault(a, "ShipStation")
    Check suggestion("Value") = "COMMAND DEFAULT", "Command default resolves"
    SetFieldDefault "Group", g, "ShipStation", "GROUP DEFAULT"
    Set suggestion = ResolveFieldDefault(a, "ShipStation", g)
    Check suggestion("Value") = "GROUP DEFAULT", "Group overrides command default"
    SetFieldDefault "Report", a, "ShipStation", Null
    Set suggestion = ResolveFieldDefault(a, "ShipStation", g)
    Check IsNull(suggestion("Value")), "Explicit null report default overrides inherited text"
    Check IsNull(MajorScalar("SELECT ShipStation FROM Report WHERE ReportID=" & a)), "Defaults do not silently commit report fields"
    snap = FreezeSummaryGroup(g, "Synthetic refreshed roster")
    meeting = CreateComparisonSet(snap, "Synthetic proposal batch")
    proposal = ProposeMeetingField(meeting, a, "Title", "Meeting batch A")
    undo = ProposeMeetingField(meeting, b, "Title", "Meeting batch B")
    ApplyMeetingProposals proposal & "," & undo
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & a) = "Meeting batch A", "Meeting batch changes first report"
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & b) = "Meeting batch B", "Meeting batch changes second report atomically"
    SetMeetingOrder meeting, b & "," & a
    Check MajorScalar("SELECT SortOrdinal FROM ComparisonMember WHERE SetID=" & meeting & " AND ReportID=" & b) = 1, "Manual meeting order retained"
    Check MajorScalar("SELECT Title FROM Report WHERE ReportID=" & b) = "Meeting batch B", "Display ordering does not assign recommendations or alter draft values"
    firstProfile = MajorAuthor(): secondProfile = AddContributor("Synthetic second reviewer")
    template = CreateRouteTemplate("Two reviewers", firstProfile & "," & secondProfile, "Independent steps")
    route = StartRoute(b, template)
    DecideRouteStep route, "Accepted", "First step complete"
    Check MajorScalar("SELECT CurrentStep FROM RouteRun WHERE RunID=" & route) = 2, "Acceptance advances to second reviewer"
    Check MajorScalar("SELECT RunState FROM RouteRun WHERE RunID=" & route) = "Pending", "Intermediate acceptance does not approve route"
    On Error Resume Next
    DecideRouteStep route, "Accepted", "Wrong reviewer"
    failed = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Check failed, "Wrong contributor cannot decide second step"
    Call SwitchContributor(secondProfile)
    DecideRouteStep route, "Returned", "Needs correction"
    Check MajorScalar("SELECT RunState FROM RouteRun WHERE RunID=" & route) = "Returned", "Assigned second reviewer can return route"
    Call SwitchContributor(firstProfile)
    template = CreateRouteTemplate("Repeated reviewer", firstProfile & "," & firstProfile, "Two distinct steps")
    Check DCount("*", "RouteTemplateStep", "TemplateID=" & template) = 2, "Same reviewer can appear at different steps"
    route = StartRoute(b, template, g)
    DecideRouteStep route, "Accepted", "First repeated step"
    DecideRouteStep route, "Accepted", "Final repeated step"
    siblingRoute = StartRoute(a, template, g)
    DecideRouteStep siblingRoute, "Accepted", "Sibling first step"
    DecideRouteStep siblingRoute, "Accepted", "Sibling final step"
    ApproveSavedRoute route, "Synthetic local approval"
    Check ApprovalIsCurrent(route), "Local approval is current after locking"
    Check RouteIsCurrent(siblingRoute), "Sibling lock incorrectly stales reviewed content"
    ApproveSavedRoute siblingRoute, "Synthetic sibling approval"
    Check ApprovalIsCurrent(route) And ApprovalIsCurrent(siblingRoute), "Approving another member stales prior approval"
    Check MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & b) = "Approved", "Local approval locks report editing"
    rev = CurrentRevision(b)
    On Error Resume Next
    Call CommitDraft(b, "EVAL", "Forbidden locked edit", "Example B", Null, Null, "", "Locked test")
    failed = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    Check failed And CurrentRevision(b) = rev, "Locked edit rejected without history"
    RemoveGroupMember g, a
    Check Not ApprovalIsCurrent(route), "Membership change makes local approval stale"
    ReopenApprovedReport b, "Synthetic correction"
    Check MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & b) = "Draft", "Explicit reopen restores editing"
    Check MajorScalar("SELECT RunState FROM RouteRun WHERE RunID=" & route) = "Reopened", "Reopening retains approval history"
    Check DCount("*", "ReportEvent", "ReportID=" & b) = 2, "Approval and reopening have contributor events"
    TestMajor = "PASS: " & count & " major workflow assertions"
    Exit Function
Failed:
    TestMajor = "FAIL after " & count & " assertions: " & Err.Number & ": " & Err.Description
End Function

'Use the installed version's importer: <=32 retains DN text as provenance;
'>=33 maps the same native values into the new independent authored fields.
Public Function SeedChiefTransfer(ByVal nativePath As String, ByVal exported As String, ByVal template As String) As String
    Dim db As DAO.Database, rows As DAO.Recordset, external As Long, id As Long
    On Error GoTo Failed
    BindActiveContributor CLng(DMin("UserProfileID", "UserProfile"))
    Set db = DBEngine.OpenDatabase(nativePath, False, True)
    Set rows = db.OpenRecordset("SELECT ReportID FROM Reports WHERE ReportType='Chief'", dbOpenSnapshot)
    external = rows!ReportID: rows.Close: db.Close
    Call OpenReviewFile(nativePath)
    id = ImportReviewedReport(external, SourceFileHash(nativePath), 2000, "SYNTHETIC TRANSFER CHIEF")
    CloseReviewFile
    Call NAVFITExportReport(id, exported, template)
    SeedChiefTransfer = "PASS: synthetic Chief transfer report " & id & "; original native comment columns retained": Exit Function
Failed:
    SeedChiefTransfer = "FAIL: " & Err.Number & " " & Err.Description
End Function

Public Function VerifyChiefTransfer(ByVal sourceSchema As Long, ByVal nativePath As String, ByVal oldExport As String, ByVal exported As String, ByVal template As String) As String
    Dim id As Long, block As Long, entry As Long, revision As Long, key As String, column As String, prefixes As Variant, expected As Object, fields As Object, changes As Object
    Dim db As DAO.Database, rows As DAO.Recordset, preview As Collection, returned As String, started As Long
    On Error GoTo Failed
    started = count: BindActiveContributor CLng(DMin("UserProfileID", "UserProfile"))
    id = DMax("ReportID", "Report", "FormKind='CHIEFEVAL'"): revision = CurrentRevision(id)
    prefixes = Array("LEAD", "TAC", "PROF", "MIS", "EO", "TEAM", "MIL")
    Set expected = CreateObject("Scripting.Dictionary")
    Set db = DBEngine.OpenDatabase(nativePath, False, True): Set rows = db.OpenRecordset("SELECT * FROM Reports WHERE ReportType='Chief'", dbOpenSnapshot)
    Set fields = SavedFields(id)
    For block = 33 To 39
        For entry = 1 To 3
            key = "Trait" & block & "Comment" & entry: column = prefixes(block - 33) & "DN" & entry
            expected(key) = rows.Fields(column).Value
            If sourceSchema < 33 Then
                Check IsNull(fields(key)), "Migration fabricated an authored legacy comment"
                Check RevisionFieldStatus(revision, revision, key) = "Not recorded", "Older comment snapshot was treated as corruption or invented history"
            Else
                Check SameFieldValue(fields(key), expected(key)), "Migration lost an authored comment"
            End If
        Next
    Next
    rows.Close: db.Close
    Call NAVFITExportReport(id, exported, template)
    Set db = DBEngine.OpenDatabase(exported, False, True): Set rows = db.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
    For block = 33 To 39
        For entry = 1 To 3
            key = "Trait" & block & "Comment" & entry: column = prefixes(block - 33) & "DN" & entry
            Check SameFieldValue(rows.Fields(column).Value, expected(key)), "Export erased untouched native comment: " & key
        Next
    Next
    rows.Close: db.Close
    Set preview = NAVFITPreviewReturned(exported, exported & ".swift-baseline.accdb")
    For block = 33 To 39
        For entry = 1 To 3
            key = "Trait" & block & "Comment" & entry
            Check NAVFITReturnedDecision(key) = "Keep current", "Unchanged exported comment became an edit: " & key
        Next
    Next
    NAVFITClearReturned
    Set preview = NAVFITPreviewReturned(oldExport, oldExport & ".swift-baseline.accdb")
    Check CurrentRevision(id) = revision, "Old baseline preview created a revision"
    For block = 33 To 39
        For entry = 1 To 3
            key = "Trait" & block & "Comment" & entry
            Check NAVFITReturnedDecision(key) = "Keep current", "Older sealed baseline lost native field identity: " & key
        Next
    Next
    NAVFITClearReturned
    returned = exported & ".returned.accdb": FileCopy exported, returned
    Set db = DBEngine.OpenDatabase(returned): db.Execute "UPDATE Reports SET LEADDN1='SYNTHETIC LEGACY RETURN' WHERE ReportType Is Not Null", dbFailOnError: db.Close
    Set preview = NAVFITPreviewReturned(returned, exported & ".swift-baseline.accdb")
    If sourceSchema < 33 Then Check NAVFITReturnedDecision("Trait33Comment1") = "Conflict", "Legacy unknown/current blank did not require a return decision"
    NAVFITResolveReturned "Trait33Comment1", "Accept returned"
    Check NAVFITApplyReturned() = id, "Accepted comment return failed"
    Set changes = CreateObject("Scripting.Dictionary"): changes("Title") = "Synthetic unrelated title change"
    Call CommitSavedFields(id, CurrentRevision(id), changes, "Unrelated edit must preserve unknown native comments")
    Call NAVFITExportReport(id, exported & ".again.accdb", template)
    Set db = DBEngine.OpenDatabase(exported & ".again.accdb", False, True): Set rows = db.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
    For block = 33 To 39
        For entry = 1 To 3
            key = "Trait" & block & "Comment" & entry: column = prefixes(block - 33) & "DN" & entry
            If key = "Trait33Comment1" Then expected(key) = "SYNTHETIC LEGACY RETURN"
            Check SameFieldValue(rows.Fields(column).Value, expected(key)), "Unrelated save erased retained comment: " & key
        Next
    Next
    rows.Close: db.Close
    VerifyChiefTransfer = "PASS: " & count - started & " Chief schema " & sourceSchema & " migration/native-provenance/old-baseline assertions": Exit Function
Failed:
    VerifyChiefTransfer = "FAIL: " & Err.Number & " " & Err.Source & " " & Err.Description
End Function

' Both writable names address the original stored ID, never RevisionNumber.
Public Function TestMajorRevisionAliases(Optional ByVal trace As Boolean = False) As String
    Dim db As DAO.Database, rows As DAO.Recordset, stored As DAO.Recordset, w As DAO.Workspace, trans As Boolean
    Dim spec As Variant, parts As Variant, id As Long, original As Long, alternate As Long, checks As Long
    On Error GoTo Failed
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0)
    For Each spec In Array("ChangeItem|qryChangeItemRevisionIDs|ItemID|ExpectedRevision|ExpectedRevisionID", "MeetingProposal|qryMeetingProposalRevisionIDs|ProposalID|BaseRevision|BaseRevisionID")
        parts = Split(CStr(spec), "|")
        If trace Then MajorAliasTrace parts(0) & ": read"
        Set rows = db.OpenRecordset(CStr(parts(1)), dbOpenDynaset)
        If rows.EOF Or Not rows.Updatable Then Err.Raise 5, "MajorHarness.TestMajorRevisionAliases", "Populated writable alias required: " & parts(1)
        Do While Not rows.EOF
            If rows.Fields(CStr(parts(4))).Value <> DLookup(CStr(parts(3)), CStr(parts(0)), parts(2) & "=" & rows.Fields(CStr(parts(2))).Value) Then Err.Raise 5, , "Alias changed stored revision ID"
            If DCount("*", "ReportRevision", "RevisionID=" & rows.Fields(CStr(parts(4))).Value) <> 1 Then Err.Raise 5, , "Alias is not a RevisionID"
            checks = checks + 2: rows.MoveNext
        Loop
        rows.MoveFirst: id = rows.Fields(CStr(parts(2))).Value: original = rows.Fields(CStr(parts(4))).Value
        alternate = DMin("RevisionID", "ReportRevision", "RevisionID<>" & original)
        If trace Then MajorAliasTrace parts(0) & ": begin transaction"
        w.BeginTrans: trans = True
        If trace Then MajorAliasTrace parts(0) & ": alias write"
        rows.Edit: rows.Fields(CStr(parts(4))) = alternate: rows.Update
        If trace Then MajorAliasTrace parts(0) & ": read stored ID"
        Set stored = db.OpenRecordset("SELECT [" & parts(3) & "] FROM [" & parts(0) & "] WHERE [" & parts(2) & "]=" & id, dbOpenSnapshot)
        If stored.Fields(0) <> alternate Then Err.Raise 5, , "Alias write did not reach legacy storage"
        stored.Close
        If trace Then MajorAliasTrace parts(0) & ": legacy write"
        db.Execute "UPDATE [" & parts(0) & "] SET [" & parts(3) & "]=" & original & " WHERE [" & parts(2) & "]=" & id, dbFailOnError
        If trace Then MajorAliasTrace parts(0) & ": rollback"
        w.Rollback: trans = False
        If trace Then MajorAliasTrace parts(0) & ": close rolled-back cursor"
        rows.Close
        If trace Then MajorAliasTrace parts(0) & ": read restored ID"
        If DLookup(CStr(parts(4)), CStr(parts(1)), parts(2) & "=" & id) <> original Then Err.Raise 5, , "Legacy write/rollback not visible through alias"
        If trace Then MajorAliasTrace parts(0) & ": FK rejection"
        If Not MajorSQLRejected("UPDATE [" & parts(1) & "] SET [" & parts(4) & "]=-1 WHERE [" & parts(2) & "]=" & id) Then Err.Raise 5, , "Alias bypassed revision foreign key"
        If DLookup(CStr(parts(4)), CStr(parts(1)), parts(2) & "=" & id) <> original Then Err.Raise 5, , "Rejected alias write changed original ID"
        checks = checks + 4
    Next
    Set rows = db.OpenRecordset("SELECT TOP 1 RevisionID FROM ReportRevision WHERE RevisionID<>RevisionNumber", dbOpenSnapshot)
    If rows.EOF Then Err.Raise 5, , "Test must distinguish IDs from revision numbers"
    rows.Close
    TestMajorRevisionAliases = "PASS: " & checks + 1 & " revision-ID alias, rollback and FK assertions": Exit Function
Failed:
    Dim primary As String
    primary = Err.Number & " " & Err.Source & " " & Err.Description
    On Error Resume Next
    If trace Then MajorAliasTrace "FAIL: " & primary
    If Not stored Is Nothing Then stored.Close
    If Not rows Is Nothing Then rows.Close
    If trans Then w.Rollback
    On Error GoTo 0
    TestMajorRevisionAliases = "FAIL: " & primary
End Function

Private Sub MajorAliasTrace(ByVal stage As String)
    Dim file As Integer
    file = FreeFile
    Open CurrentProject.Path & "\major-alias-trace.log" For Append As #file
    Print #file, stage
    Close #file
End Sub

Private Function MajorSQLRejected(ByVal sql As String) As Boolean
    On Error GoTo Rejected
    CurrentDb.Execute sql, dbFailOnError
    Exit Function
Rejected:
    MajorSQLRejected = True
End Function

' Exercise the current comparison and the retained hidden workbench adapter;
' opening the retired visible workbench route is not a supported test setup.
Public Function TestMajorAliasViews() As String
    Dim groupID As Long, reportID As Long, snapshot As Long, meeting As Long, proposal As Long, batch As Long, revision As Long, header As Long, f As Access.Form, changes As Object
    On Error GoTo Failed
    groupID = DMin("GroupID", "SummaryGroup"): reportID = DMin("ReportID", "GroupMember", "GroupID=" & groupID)
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "Trait33", "4.0"
    Call CommitSavedFields(reportID, CurrentRevision(reportID), changes, "Synthetic alias before value")
    revision = CurrentRevision(reportID)
    snapshot = FreezeSummaryGroup(groupID, "Synthetic alias view check")
    meeting = CreateComparisonSet(snapshot, "Synthetic alias view")
    proposal = ProposeMeetingField(meeting, reportID, "Trait33", "3.0")
    DoCmd.OpenForm "frmUXComparison", , , , , , CStr(meeting)
    Set f = Forms!frmUXComparison
    Check UXRefreshComparisonPreview(), "Comparison alias preview failed"
    header = Abs(f!lstProposals.ColumnHeads)
    Check f!lstProposals.ListCount - header = 1, "Proposal alias list omitted row"
    Check SameFieldValue(f!lstProposals.Column(3, header), MajorRevisionValue(revision, "Trait33")), "Proposal before value differs: " & Nz(f!lstProposals.Column(3, header), "[null]") & "; row zero=" & Nz(f!lstProposals.Column(3, 0), "[null]")
    Check f!lstProposals.Column(4, header) = "3.0", "Proposal after value differs"
    f!lstProposals = proposal
    Check UXGridDetail(), "Proposal alias detail failed"
    Check InStr(f!uxProposalDetail, "AFTER: 3.0") > 0, "Full proposal detail missing"
    Check UXCloseComparison(False), "Comparison close failed"
    DoCmd.OpenForm "frmMajorWorkbench", , , , , acHidden, "ContextOnly"
    Check LoadMajorWorkbench(), "Hidden compatibility adapter failed"
    Set f = Forms!frmMajorWorkbench
    batch = PreviewFieldBatch(CStr(reportID), "Title", "Synthetic alias batch", "Synthetic view check")
    f!cboBatch = batch
    Check MajorAction("ShowBatch") And f!lstDetails.ListCount - Abs(f!lstDetails.ColumnHeads) = 1, "Batch alias list failed"
    header = Abs(f!lstDetails.ColumnHeads): f!lstDetails.Selected(header) = True
    Check ShowMajorDetails() And InStr(f!txtDetail, "Synthetic alias batch") > 0, "Batch alias detail failed"
    f!cboMeeting = meeting
    Check MajorAction("ShowProposals") And f!lstDetails.ListCount - Abs(f!lstDetails.ColumnHeads) = 1, "Workbench proposal alias list failed"
    Check CLng(f!lstDetails.Column(5, Abs(f!lstDetails.ColumnHeads))) = revision, "Workbench proposal lost base RevisionID"
    Check CurrentRevision(reportID) = revision, "View preview changed saved report"
    DoCmd.Close acForm, "frmMajorWorkbench", acSaveNo
    TestMajorAliasViews = "PASS: 13 current comparison/compatibility-adapter alias assertions": Exit Function
Failed:
    TestMajorAliasViews = "FAIL: " & Err.Number & " " & Err.Source & " " & Err.Description
End Function
