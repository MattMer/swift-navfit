Attribute VB_Name = "SignoffHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 9690, , message
    checks = checks + 1
End Sub
Private Function MarkFails(ByVal id As Long, ByVal revision As Long, ByVal signedOn As Date) As Boolean
    On Error GoTo Failed
    Call MarkReportSigned(id, revision, signedOn)
    Exit Function
Failed:
    MarkFails = True
End Function
Private Function MemberChangeFails(ByVal groupID As Long, ByVal reportID As Long, ByVal adding As Boolean) As Boolean
    On Error GoTo Failed
    If adding Then AddGroupMember groupID, reportID, "E5" Else RemoveGroupMember groupID, reportID
    Exit Function
Failed:
    MemberChangeFails = InStr(Err.Description, "Reopen the completed group") > 0
End Function
Private Function SaveFails(ByVal id As Long) As Boolean
    Dim changes As Object
    On Error GoTo Failed
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "OfficialComments", "Changed while signed"
    Call CommitSavedFields(id, CurrentRevision(id), changes, "Should be refused")
    Exit Function
Failed:
    SaveFails = True
End Function
Private Function RestoreFails(ByVal id As Long, ByVal revision As Long) As Boolean
    On Error GoTo Failed
    Call RestoreDraft(id, revision)
    Exit Function
Failed:
    RestoreFails = True
End Function
Private Function CompareFails(ByVal id As Long, ByVal older As Long) As Boolean
    On Error GoTo Failed
    Call OpenRevisionComparison(id, older)
    Exit Function
Failed:
    CompareFails = True
End Function
Private Function Revise(ByVal id As Long, ByVal text As String) As Long
    Dim changes As Object
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "OfficialComments", text
    Revise = CommitSavedFields(id, CurrentRevision(id), changes, "Draft: " & text)
    Revise = CurrentRevision(id)
End Function
Public Function TestSignoff() As String
    Dim id As Long, other As Long, groupID As Long, draft1 As Long, draft2 As Long, signed1 As Long, signed2 As Long, reopened As Long
    Dim stage As String, plan As Object, f As Access.Form, rows As Access.Form, ops As Long, receipt As Long, message As Long, before As Object, after As Object, key As Variant
    Dim fields As Object, batch As Long, s As Object, disposal As Long, found As Boolean, i As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC SIGNOFF", "Signoff tester"

    stage = "setup"
    id = CommitDraft(0, "EVAL", "Signoff", "SYNTHETIC, SIGNED", #9/16/2025#, #9/15/2026#, "", "Test", "First draft")
    draft1 = CurrentRevision(id)
    draft2 = Revise(id, "Second draft wording")
    message = AddReportMessage(id, "Reviewer note kept with the report")
    receipt = ImportReceiptFor(id)
    CurrentDb.Execute "INSERT INTO ReportNoteDraft (DraftGuid,ReportID,AuthorProfileID,NoteText,VersionCounter,CheckpointAt) VALUES ('" & NewIdentity() & "'," & id & "," & MajorAuthor() & ",'Unposted note',1,Now())", dbFailOnError
    other = DuplicateDraft(id)
    groupID = CreateSummaryGroup("Signoff group", "EVAL", "E5"): AddGroupMember groupID, id, "E5": AddGroupMember groupID, other, "E5"

    stage = "mark rules"
    Check MarkFails(id, draft1, #10/1/2026#), "Stale revision accepted"
    Check MarkFails(id, draft2, Date + 1), "Future signed date accepted"
    Check MarkFails(999999, draft2, #10/1/2026#), "Missing report accepted"
    ops = DCount("*", "OperationJournal", "OperationKind='Mark report signed'")
    signed1 = draft2
    Check MarkReportSigned(id, signed1, #10/1/2026#) > 0, "Mark signed failed"
    Check ReportIsSigned(id) And DLookup("WorkflowState", "Report", "ReportID=" & id) = "Signed", "Workflow state not Signed"
    Check DLookup("SignedRevisionID", "ReportSignoff", "ReportID=" & id) = signed1 And DLookup("PriorState", "ReportSignoff", "ReportID=" & id) = "Draft", "Signoff row wrong"
    Check DCount("*", "ReportEvent", "ReportID=" & id & " AND EventKind='Marked signed'") = 1, "Signed event missing"
    Check DCount("*", "AuditEvent", "ReportEventSourceID IN (SELECT ReportEventID FROM ReportEvent WHERE ReportID=" & id & " AND EventKind='Marked signed')") = 1, "Signed event not in audit index"
    Check DCount("*", "OperationJournal", "OperationKind='Mark report signed'") = ops + 1, "Mark signed not journaled"
    Check MarkFails(id, signed1, #10/1/2026#), "Second mark accepted"
    Check UXReportState(id) = "Signed", "Report list state not Signed"

    stage = "locked routes"
    Check SaveFails(id), "Save to a signed report accepted"
    Check RestoreFails(id, draft1), "Restore into a signed report accepted"
    Set fields = CreateObject("Scripting.Dictionary"): fields.Add "OfficialComments", "Batch change"
    On Error Resume Next
    batch = PreviewReportChanges(CStr(id), fields, "Signed batch")
    If Err.Number = 0 Then ApplyFieldBatch batch
    Check Err.Number <> 0, "Field batch changed a signed report": Err.Clear
    On Error GoTo Failed
    Check ReportIsSigned(id) And CurrentRevision(id) = signed1, "Refused routes changed the report"
    Set plan = RecursiveDefaultPlan(groupID, False, id & "," & other, "UnitUIC", "00123")
    Check plan("Reports").Count = 1 And plan("SignedSkipped") = 1, "Recursive defaults did not skip the signed report"

    stage = "editor prompt"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Set f = Forms!frmEditor
    Check Not f!cmdSave.Enabled And InStr(f!uxLocked.Caption, "Marked signed") = 1, "Signed editor state wrong: " & f!uxLocked.Caption
    Check Not f!uxMarkSigned.Enabled, "Mark signed offered on a signed report"
    UXSignedTestAnswer vbNo
    Check Not UXSignedEditGuard(), "No did not refuse the edit"
    Check ReportIsSigned(id), "No reopened the report"
    UXSignedTestAnswer vbNo
    Forms!frmEditor!uxFieldSurface.Form.Controls(UXFieldControl("OccasionPeriodic")).Value = True
    Call UXSurfaceEdited("OccasionPeriodic")
    Check Nz(Forms!frmEditor.Controls(UXFieldControl("OccasionPeriodic")), "") <> "Yes" And Not EditorHasPending(), "Refused checkbox edit reached the buffer"
    UXSignedTestAnswer vbYes, "Corrected member name"
    Check UXSurfaceSetValue("OfficialName", "SYNTHETIC, CORRECTED"), "Edit after reopening failed"
    Check Not ReportIsSigned(id) And DLookup("WorkflowState", "Report", "ReportID=" & id) = "Draft", "Yes did not reopen"
    Check DLookup("ReopenReason", "ReportSignoff", "ReportID=" & id) = "Corrected member name" And Not IsNull(DLookup("ReopenedAt", "ReportSignoff", "ReportID=" & id)), "Reopen not recorded"
    Check DCount("*", "ReportEvent", "ReportID=" & id & " AND EventKind='Reopened for correction'") = 1, "Reopen event missing"
    Check f!cmdSave.Enabled And EditorHasPending() And UXReportState(id) = "Reopened for correction", "Reopened editor state wrong"
    Check UXSaveEditor(), "Save after reopening failed: " & f!uxError.Caption & " | " & f!lblSaved.Caption
    signed2 = CurrentRevision(id)
    Check signed2 <> signed1, "Correction did not create a revision"

    stage = "re-sign with disposal"
    Check UXSignoffOpen(), "Mark signed dialog did not open"
    Forms!frmUXSignoff!txtSignedOn = #10/5/2026#: Forms!frmUXSignoff!chkDispose = True
    If Not UXSignoffSave(True) Then Check False, "Re-sign with disposal failed: " & Forms!frmUXSignoff!lblStatus.Caption
    Check Not CurrentProject.AllForms("frmUXSignoff").IsLoaded, "Mark signed dialog stayed open"
    Check ReportIsSigned(id) And DCount("*", "ReportSignoff", "ReportID=" & id) = 2, "Second signoff missing"
    disposal = Nz(DMax("DisposalID", "DraftDisposal", "ReportID=" & id), 0)
    Check disposal > 0 And DLookup("RevisionsCleared", "DraftDisposal", "DisposalID=" & disposal) = 1, "Disposal record wrong: " & DLookup("RevisionsCleared", "DraftDisposal", "DisposalID=" & disposal)
    Check DLookup("ValuesDisposed", "ReportRevision", "RevisionID=" & draft1) = True And DCount("*", "RevisionValue", "RevisionID=" & draft1) = 0, "Draft revision not emptied"
    Check DCount("*", "RevisionValue", "RevisionID=" & signed1) > 0 And DCount("*", "RevisionValue", "RevisionID=" & signed2) > 0, "A signed revision lost values"
    Check Nz(DLookup("ValuesDisposed", "ReportRevision", "RevisionID=" & signed1), False) = False, "Signed revision marked disposed"
    Check DCount("*", "RevisionChange", "RevisionID=" & signed1) = 0, "Draft values kept in the first signed version's change list"
    Check DCount("*", "RevisionChange", "RevisionID=" & signed2) > 0, "Change between two signed versions removed"
    Check SavedFields(id)("OfficialName") = "SYNTHETIC, CORRECTED", "Current values changed by disposal"
    Check DCount("*", "ReportMessage", "MessageID=" & message) = 1, "Conversation removed"
    Check DCount("*", "ReportNoteDraft", "ReportID=" & id) = 0, "Unposted note draft kept"
    Check DCount("*", "ImportRawValue", "ImportReceiptID=" & receipt) = 0 And DCount("*", "ImportReceipt", "ImportReceiptID=" & receipt) = 1, "Raw values kept or receipt removed"
    Check DCount("*", "GroupMember", "ReportID=" & id) = 1, "Group membership changed"
    Check RestoreFails(id, draft1) And CompareFails(id, draft1), "Disposed draft could be restored or compared"
    Call UXSelectSection(9)
    For i = 0 To f!lstRevisions.ListCount - 1
        If InStr(Nz(f!lstRevisions.Column(3, i), ""), "[Draft contents disposed]") = 1 Then found = True
    Next
    Check found And f!lstRevisions.ListCount - Abs(f!lstRevisions.ColumnHeads) = DCount("*", "ReportRevision", "ReportID=" & id & " AND " & HistoryRevisionPredicate("ReportRevision")), "History list missing retained checkpoints or the disposed marker"
    Check DCount("*", "ReportEvent", "ReportID=" & id & " AND EventKind='Drafts disposed'") = 1, "Disposal event missing"
    Check Not UXDisposeDrafts(id, True), "Second disposal found drafts"
    Call UXLeaveEditor("Discard")

    stage = "group grid"
    Check MarkReportSigned(other, CurrentRevision(other), #10/2/2026#) > 0, "Mark other signed failed"
    Check UXShowGroupMembers(groupID), "Members did not open": Set rows = Forms!frmUXComparison!uxTraitGrid.Form
    rows.Recordset.FindFirst "ReportID=" & other
    Check Not rows.Recordset.NoMatch, "Signed member row missing"
    rows!Trait33 = "3.0"
    Check Not UXGridEdit("Trait33", vbNo) And ReportIsSigned(other), "Grid No changed the signed report"
    rows.Recordset.FindFirst "ReportID=" & other
    rows!Trait33 = "3.0"
    Check UXGridEdit("Trait33", vbYes), "Grid Yes did not save: " & Forms!frmUXComparison!lblStatus.Caption
    Check Not ReportIsSigned(other) And SavedFields(other)("Trait33") = "3.0", "Grid reopen did not save"
    DoCmd.Close acForm, "frmUXComparison"

    stage = "group signing"
    Dim g As Long, r1 As Long, r2 As Long, r3 As Long, senior As Object, sid As Long
    'Completion posts RSCA accounting, so the reports carry a senior and marks.
    Set senior = CreateObject("Scripting.Dictionary")
    For Each key In RSSeniorKeys(): senior.Add CStr(key), Null: Next
    senior("ReportingSeniorSSN") = "000000001": senior("ReportingSeniorName") = "SYNTHETIC"
    sid = RSSaveSenior(0, "Synthetic senior", senior)
    r1 = AccountedReport(senior, "Group sign one", "IT2"): r2 = AccountedReport(senior, "Group sign two", "")
    g = CreateSummaryGroup("Group signing", "EVAL", "E5"): AddGroupMember g, r1, "E5": AddGroupMember g, r2, "E5"
    RSSetGroupBaseline g, sid, RSAddBaseline(sid, "E5", #12/31/2025#, 4.2, 100, 420, "Synthetic baseline"), True
    Check MarkGroupReportsSigned(g, #10/9/2026#) = 2 And ReportIsSigned(r1) And ReportIsSigned(r2), "Mark all reports signed failed"
    Check MarkGroupReportsSigned(g, #10/9/2026#) = 0, "Already signed reports marked again"
    ' Reports signed one by one were not signed by a completion: a group reopen leaves them.
    Check ReopenGroupReports(g, "Test reopen") = 0 And ReportIsSigned(r1) And ReportIsSigned(r2), "Group reopen undid signatures the group did not make"
    Check ReopenSignedReport(r1, "Test reopen") And ReopenSignedReport(r2, "Test reopen") And Not ReportIsSigned(r1), "Individual reopen failed"
    SetGroupWorkflow g, GroupStamp(g), "Completed", #10/10/2026#
    Check ReportIsSigned(r1) And ReportIsSigned(r2) And DLookup("SignedOn", "ReportSignoff", "SignoffID=" & ReportCurrentSignoff(r2)) = #10/10/2026#, "Completing the group did not sign its reports on the given date"
    SetGroupWorkflow g, GroupStamp(g), "In-Work"
    Check Not ReportIsSigned(r1) And Not ReportIsSigned(r2), "Reopening the completed group left reports signed"
    ' A report signed on its own before completion keeps its signature when the group reopens.
    Call MarkReportSigned(r1, CurrentRevision(r1), #10/9/2026#)
    SetGroupWorkflow g, GroupStamp(g), "Completed", #10/10/2026#
    Check ReportIsSigned(r2) And DLookup("SignedOn", "ReportSignoff", "SignoffID=" & ReportCurrentSignoff(r1)) = #10/9/2026#, "Completion re-signed a report already signed"
    SetGroupWorkflow g, GroupStamp(g), "In-Work"
    Check ReportIsSigned(r1) And Not ReportIsSigned(r2), "Group reopen reopened a report signed separately, or kept its own signature"
    Check ReopenSignedReport(r1, "Test reopen"), "Separately signed report could not be reopened"
    SetGroupWorkflow g, GroupStamp(g), "Completed"
    Check Not ReportIsSigned(r1), "Completing without a signed date signed reports"
    'A Completed group's members change only after Reopen, and completing again
    'with a signed date signs every member still unsigned.
    r3 = AccountedReport(senior, "Group sign three", "")
    Check MemberChangeFails(g, r3, True) And DCount("*", "GroupMember", "GroupID=" & g) = 2, "Completed group accepted a new member"
    Check MemberChangeFails(g, r2, False) And DCount("*", "GroupMember", "GroupID=" & g) = 2, "Completed group lost a member"
    SetGroupWorkflow g, GroupStamp(g), "Completed", #10/10/2026#
    Check ReportIsSigned(r1) And ReportIsSigned(r2), "Completing again did not sign the unsigned members"
    SetGroupWorkflow g, GroupStamp(g), "In-Work"
    Check UXSignoffOpenGroup(g, False), "Group dialog did not open"
    Forms!frmUXSignoff!txtSignedOn = #10/11/2026#
    If Not UXSignoffSave(True) Then Check False, "Group dialog save failed: " & Forms!frmUXSignoff!lblStatus.Caption
    Check ReportIsSigned(r1) And ReportIsSigned(r2) And GroupWorkflow(g) = "In-Work", "Group dialog did not sign reports, or changed the group status"

    stage = "disposal failure after signing"
    ' Signing commits first; a failed optional disposal is reported on its own and
    ' can be retried through Dispose of drafts.
    Dim partial As Long
    partial = CommitDraft(0, "EVAL", "Disposal failure", "SYNTHETIC, DISPOSAL", #1/1/2026#, #9/30/2026#, "")
    Call Revise(partial, "Second saved revision")
    If CurrentProject.AllForms("frmEditor").IsLoaded Then DoCmd.Close acForm, "frmEditor", acSaveNo
    Check UXSignoffOpen(partial), "Mark signed dialog did not open for the disposal case"
    CurrentDb.TableDefs("DraftDisposal").Fields("RevisionsCleared").ValidationRule = "<0"
    Forms!frmUXSignoff!txtSignedOn = #10/9/2026#: Forms!frmUXSignoff!chkDispose = True
    Check UXSignoffSave(True), "Signing reported failure although it committed"
    CurrentDb.TableDefs("DraftDisposal").Fields("RevisionsCleared").ValidationRule = ""
    Check ReportIsSigned(partial) And DCount("*", "DraftDisposal", "ReportID=" & partial) = 0, "Signing or disposal outcome wrong"
    Check CurrentProject.AllForms("frmUXSignoff").IsLoaded, "Dialog closed and hid the disposal failure"
    Check InStr(Forms!frmUXSignoff!lblStatus.Caption, "Marked signed. Disposing of drafts failed") = 1 And Not Forms!frmUXSignoff!cmdSave.Enabled, "Disposal failure message or retry guard wrong: " & Forms!frmUXSignoff!lblStatus.Caption
    DoCmd.Close acForm, "frmUXSignoff"
    Check UXDisposeDrafts(partial, True) And DCount("*", "DraftDisposal", "ReportID=" & partial) = 1, "Disposal retry failed"

    stage = "lifecycle and signing ends review"
    ' One lifecycle answer; signing a report under review ends the route or assignment.
    Dim routed As Long, assigned As Long, reviewer As Long, requester As Long, template As Long, run As Long, task As Long
    requester = ActiveContributorID(): reviewer = AddContributor("Signoff reviewer")
    template = CreateRouteTemplate("Signoff route", CStr(reviewer), "")
    routed = CommitDraft(0, "EVAL", "Routed", "SYNTHETIC, ROUTED", #9/16/2025#, #9/15/2026#, "", "Test")
    assigned = CommitDraft(0, "EVAL", "Assigned", "SYNTHETIC, ASSIGNED", #9/16/2025#, #9/15/2026#, "", "Test")
    Check ReportLifecycle(routed) = "Draft" And ReportLifecycle(0) = "Missing", "Draft or missing lifecycle"
    run = StartRoute(routed, template)
    task = AssignReview(assigned, CurrentRevision(assigned), reviewer, Null, "Check it")
    Check ReportLifecycle(routed) = "In review" And ReportLifecycle(assigned) = "In review", "Open review not reported by the lifecycle"
    Check Not ReportTakesGroupChange("In review") And ReportTakesGroupChange("Approved") And ReportTakesGroupChange("Draft"), "Group change rule"
    Call UXSignoffOpen(routed)
    Check InStr(Forms!frmUXSignoff!lblStatus.Caption, "Marking it signed ends that review") > 0, "Sign-off window did not warn about the open review"
    DoCmd.Close acForm, "frmUXSignoff"
    Call MarkReportSigned(routed, CurrentRevision(routed), #9/20/2026#)
    Call MarkReportSigned(assigned, CurrentRevision(assigned), #9/20/2026#)
    Check DLookup("RunState", "RouteRun", "RunID=" & run) = "Cancelled" And DCount("*", "RouteRunEvent", "RunID=" & run & " AND EventKind='Cancelled' AND Notes Like '*marked signed*'") = 1, "Signing left the route open or unexplained"
    Check DLookup("TaskState", "ReviewTask", "TaskID=" & task) = "Cancelled" And DCount("*", "ReviewTaskEvent", "TaskID=" & task & " AND EventKind='Cancelled'") = 1, "Signing left the assignment open or unexplained"
    Check ReviewHolder(routed) = 0 And ReviewHolder(assigned) = 0 And ReportLifecycle(routed) = "Signed", "A signed report is still held for review"
    Check InStr(DLookup("Outcome", "OperationJournal", "OperationID=" & DMax("OperationID", "OperationJournal", "OperationKind='Mark report signed'")), "1 open review(s) ended") > 0, "Operation record omits the ended review"
    Check ReopenSignedReport(routed, "Correction") And ReportLifecycle(routed) = "Draft", "Reopened report did not return to draft without the old review"
    Call ArchiveDraft(assigned, True)
    Check ReportLifecycle(assigned) = "Archived", "Archived lifecycle"
    ' Archiving a report under review ends the review; recovering it does not restart it.
    Dim archivedReport As Long, archivedRun As Long
    archivedReport = CommitDraft(0, "EVAL", "Archived in review", "SYNTHETIC, ARCHIVED", #9/16/2025#, #9/15/2026#, "", "Test")
    archivedRun = StartRoute(archivedReport, template)
    Check ArchiveDraft(archivedReport, True), "Archive failed"
    Check DLookup("RunState", "RouteRun", "RunID=" & archivedRun) = "Cancelled" And DCount("*", "RouteRunEvent", "RunID=" & archivedRun & " AND EventKind='Cancelled' AND Notes Like '*archived*'") = 1, "Archiving left the review open or unexplained"
    Check ReviewHolder(archivedReport) = 0 And UXReportState(archivedReport) = "Archived", "Archived report still held for review"
    Check ArchiveDraft(archivedReport, False) And ReportLifecycle(archivedReport) = "Draft" And UXReportState(archivedReport) = "Draft", "Recovered report restarted its old review"
    ' The journal moves forward only.
    Dim journal As Long, refused As Boolean
    Check OperationMayBecome("Started", "Prepared") And OperationMayBecome("Prepared", "Completed") And Not OperationMayBecome("Completed", "Failed") And Not OperationMayBecome("Started", "Started"), "Journal transition rules"
    journal = BeginOperation("Journal rule test", "Synthetic")
    FinishOperation journal, "Completed", "Done"
    On Error Resume Next
    FinishOperation journal, "Failed", "Late failure": refused = (Err.Number = vbObjectError + 1108): Err.Clear
    On Error GoTo Failed
    Check refused And DLookup("OperationState", "OperationJournal", "OperationID=" & journal) = "Completed", "A completed operation was changed"

    stage = "reports list"
    DoCmd.OpenForm "frmLibrary"
    Forms!frmLibrary!uxStatusFilter = "Signed": RefreshLibrary
    Check Forms!frmLibrary!lstReports.ListCount - Abs(Forms!frmLibrary!lstReports.ColumnHeads) = DCount("*", "Report", "WorkflowState='Signed' AND IsArchived=False"), "Signed filter wrong"
    Forms!frmLibrary!uxStatusFilter = "All": Forms!frmLibrary!cboRankFilter = "E5": RefreshLibrary
    Check Forms!frmLibrary!lstReports.ListCount - Abs(Forms!frmLibrary!lstReports.ColumnHeads) = 1 And Forms!frmLibrary!lstReports.Column(4, Abs(Forms!frmLibrary!lstReports.ColumnHeads)) = "IT2", "Rank filter or Rate column wrong"
    Call UXClearReportFilters
    Check Forms!frmLibrary!cboRankFilter = "All ranks/rates", "Clear filters kept the rank filter"
    DoCmd.Close acForm, "frmLibrary"

    TestSignoff = "PASS: " & checks & " mark signed, lock, reopen prompt, save refusal, bulk skip, draft disposal, grid, group signing and reports list assertions": Exit Function
Failed:
    TestSignoff = "FAIL at " & stage & " after " & checks & ": " & Err.Number & " " & Err.Description
End Function
Private Function AccountedReport(ByVal senior As Object, ByVal title As String, ByVal rank As String) As Long
    Dim f As Object, key As Variant
    AccountedReport = CommitDraft(0, "EVAL", title, title, #9/16/2025#, #9/15/2026#, "", "Test")
    Set f = CreateObject("Scripting.Dictionary")
    For Each key In senior.Keys: f.Add CStr(key), senior(key): Next
    f.Add "ObservationStatus", "Observed"
    For Each key In TraitKeys(): f.Add CStr(key), "4.0": Next
    If rank <> "" Then f.Add "RankRate", rank
    CommitSavedFields AccountedReport, CurrentRevision(AccountedReport), f, "Synthetic accounting fixture"
End Function
Private Function ImportReceiptFor(ByVal id As Long) As Long
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("ImportReceipt", dbOpenDynaset)
    r.AddNew: r!ReceiptGuid = NewIdentity(): r!SourceHash = String$(64, "0"): r!SourcePath = "SYNTHETIC source": r!ExternalReportID = 1: r!ReportID = id: r!AuthorProfileID = MajorAuthor(): r!ImportedAt = Now(): r!AdapterVersion = "Synthetic"
    r.Update: r.Bookmark = r.LastModified: ImportReceiptFor = r!ImportReceiptID: r.Close
    CurrentDb.Execute "INSERT INTO ImportRawValue (ImportReceiptID,SourceTable,SourceRow,FieldOrdinal,FieldKey,DAOType,RawValue,IsNullValue) VALUES (" & ImportReceiptFor & ",'Reports',1,1,'Synthetic',10,'SYNTHETIC RAW',False)", dbFailOnError
End Function
Public Function VerifySignoffRecovery() As String
    Dim id As Long
    id = Nz(DLookup("ReportID", "Report", "Title='Signoff'"), 0)
    If id = 0 Then VerifySignoffRecovery = "FAIL: report missing": Exit Function
    If Not ReportIsSigned(id) Then VerifySignoffRecovery = "FAIL: signed state lost": Exit Function
    If DCount("*", "ReportSignoff", "ReportID=" & id) <> 2 Or DCount("*", "DraftDisposal", "ReportID=" & id) <> 1 Then VerifySignoffRecovery = "FAIL: signoff history lost": Exit Function
    If SavedFields(id)("OfficialName") <> "SYNTHETIC, CORRECTED" Then VerifySignoffRecovery = "FAIL: signed values lost": Exit Function
    VerifySignoffRecovery = "PASS: signed state, signoff history, disposal record and signed values persist in a new session"
End Function
