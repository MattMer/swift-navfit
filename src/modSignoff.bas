Attribute VB_Name = "modSignoff"
Option Compare Database
Option Explicit

'Mark signed records that the official report was signed outside Swift. It is not
'an official signature or submission. A signed report is read only: the first edit
'asks whether to reopen it for correction. Draft disposal (BUPERSINST 1610.10H,
'enclosure (2) Introduction paragraph 7a: drafts destroyed when the report is signed)
'empties superseded revisions but keeps every signed revision, all conversations,
'reviews, group history and accounting, and records what it removed.
Private testDecision As Long, testReason As String, testUsed As Boolean
Private signoffReport As Long, signoffRevision As Long, signoffGroup As Long, signoffCompleting As Boolean, signoffStamp As String

Public Function ReportIsSigned(ByVal reportID As Long) As Boolean
    ReportIsSigned = (Nz(MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & reportID), "") = "Signed")
End Function
Public Function ReportCurrentSignoff(ByVal reportID As Long) As Long
    ReportCurrentSignoff = Nz(MajorScalar("SELECT TOP 1 SignoffID FROM ReportSignoff WHERE ReportID=" & reportID & " AND ReopenedAt Is Null ORDER BY SignoffID DESC"), 0)
End Function
Public Function SignedReportCaption(ByVal reportID As Long) As String
    Dim id As Long
    id = ReportCurrentSignoff(reportID)
    If id = 0 Then Exit Function
    SignedReportCaption = "Marked signed " & Format$(MajorScalar("SELECT SignedOn FROM ReportSignoff WHERE SignoffID=" & id), "d mmm yyyy") & " by " & Nz(MajorScalar("SELECT p.DisplayName FROM ReportSignoff AS s INNER JOIN UserProfile AS p ON s.MarkedBy=p.UserProfileID WHERE s.SignoffID=" & id), "unknown contributor") & ". Editing a field asks whether to reopen it for correction."
End Function

Private Sub ReportEventRecord(ByVal reportID As Long, ByVal kind As String, ByVal operation As Long)
    Dim e As DAO.Recordset
    Set e = CurrentDb.OpenRecordset("ReportEvent", dbOpenDynaset)
    e.AddNew: e!ReportID = reportID: e!AuthorProfileID = MajorAuthor(): e!EventAt = Now(): e!EventKind = kind: e.Update
    e.Bookmark = e.LastModified: AuditRecord "Report event", CLng(e!ReportEventID), operation: e.Close
End Sub

Public Function MarkReportSigned(ByVal reportID As Long, ByVal expectedRevision As Long, ByVal signedOn As Date) As Long
    Dim w As DAO.Workspace, r As DAO.Recordset, s As DAO.Recordset, state As String, operation As Long
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT ReportID,IsArchived,WorkflowState,CurrentRevisionID FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 9600, "modSignoff.MarkReportSigned", "The report no longer exists."
    If r!IsArchived Then Err.Raise vbObjectError + 9601, "modSignoff.MarkReportSigned", "Recover the archived report before marking it signed."
    state = Nz(r!WorkflowState, "")
    If state = "Signed" Then Err.Raise vbObjectError + 9602, "modSignoff.MarkReportSigned", "This report is already marked signed."
    If state <> "Draft" And state <> "Approved" Then Err.Raise vbObjectError + 9602, "modSignoff.MarkReportSigned", "Only a draft or locally approved report can be marked signed."
    If Nz(r!CurrentRevisionID, 0) = 0 Or r!CurrentRevisionID <> expectedRevision Then Err.Raise vbObjectError + 9603, "modSignoff.MarkReportSigned", "The saved report changed. Reopen it and check the signed version again."
    r.Close
    If DateValue(signedOn) > Date Then Err.Raise vbObjectError + 9604, "modSignoff.MarkReportSigned", "The signed date cannot be in the future."
    operation = BeginOperation("Mark report signed", "Report " & reportID & "; revision " & expectedRevision & "; signed " & Format$(signedOn, "yyyy-mm-dd"))
    Set s = CurrentDb.OpenRecordset("ReportSignoff", dbOpenDynaset)
    s.AddNew: s!SignoffGuid = NewIdentity(): s!ReportID = reportID: s!SignedRevisionID = expectedRevision: s!SignedOn = DateValue(signedOn)
    s!PriorState = state: s!MarkedBy = MajorAuthor(): s!MarkedAt = Now(): s.Update: s.Bookmark = s.LastModified: MarkReportSigned = s!SignoffID: s.Close
    CurrentDb.Execute "UPDATE Report SET WorkflowState='Signed' WHERE ReportID=" & reportID, dbFailOnError
    ReportEventRecord reportID, "Marked signed", operation
    FinishOperation operation, "Completed", "Report locked as signed. This records a signature made outside Swift; it is not an official signature or submission."
    w.CommitTrans: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modSignoff", "MarkReportSigned", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    If Not s Is Nothing Then SwiftClose s, swiftPrimary, "s"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function

Public Function ReopenSignedReport(ByVal reportID As Long, Optional ByVal reason As String = "") As Boolean
    Dim w As DAO.Workspace, s As DAO.Recordset, id As Long, operation As Long
    reason = Left$(Trim$(reason), 255)
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If Not ReportIsSigned(reportID) Then Err.Raise vbObjectError + 9610, "modSignoff.ReopenSignedReport", "This report is not marked signed."
    id = ReportCurrentSignoff(reportID)
    If id = 0 Then Err.Raise vbObjectError + 9611, "modSignoff.ReopenSignedReport", "The signed record for this report is missing."
    operation = BeginOperation("Reopen signed report", "Report " & reportID & IIf(reason = "", "", ": " & reason))
    Set s = CurrentDb.OpenRecordset("SELECT * FROM ReportSignoff WHERE SignoffID=" & id, dbOpenDynaset)
    s.Edit: s!ReopenedBy = MajorAuthor(): s!ReopenedAt = Now(): s!ReopenReason = IIf(reason = "", Null, reason): s.Update: s.Close
    CurrentDb.Execute "UPDATE RouteRun SET RunState='Reopened' WHERE ReportID=" & reportID & " AND RunState='Approved'", dbFailOnError
    CurrentDb.Execute "UPDATE Report SET WorkflowState='Draft' WHERE ReportID=" & reportID, dbFailOnError
    ReportEventRecord reportID, "Reopened for correction", operation
    FinishOperation operation, "Completed", "Reopened for correction. The signed revision stays in history; mark the corrected report signed again."
    w.CommitTrans: ReopenSignedReport = True: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modSignoff", "ReopenSignedReport", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    SwiftRollback w, swiftPrimary
    If Not s Is Nothing Then SwiftClose s, swiftPrimary, "s"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function

'Revisions whose contents can be disposed: not signed (now or before), not current,
'not already disposed. Returned as a comma list for IN clauses.
Public Function DisposableRevisions(ByVal reportID As Long) As String
    Dim r As DAO.Recordset, ids As String
    Set r = CurrentDb.OpenRecordset("SELECT v.RevisionID FROM ReportRevision AS v INNER JOIN Report AS p ON v.ReportID=p.ReportID WHERE v.ReportID=" & reportID & " AND v.RevisionID<>Nz(p.CurrentRevisionID,0) AND Not Nz(v.ValuesDisposed,False) AND v.RevisionID NOT IN (SELECT SignedRevisionID FROM ReportSignoff WHERE ReportID=" & reportID & ") ORDER BY v.RevisionID", dbOpenSnapshot)
    Do While Not r.EOF
        If ids <> "" Then ids = ids & ","
        ids = ids & r!RevisionID: r.MoveNext
    Loop
    r.Close
    DisposableRevisions = ids
End Function
Public Function DisposeReportDrafts(ByVal reportID As Long) As Long
    Dim w As DAO.Workspace, d As DAO.Recordset, db As DAO.Database, ids As String, kept As String, signoff As Long, operation As Long
    Dim revisions As Long, values As Long, changes As Long, recovery As Long, notes As Long, raw As Long
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If Not ReportIsSigned(reportID) Then Err.Raise vbObjectError + 9620, "modSignoff.DisposeReportDrafts", "Mark the report signed before disposing of its drafts."
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Nz(Forms!frmEditor!txtReportID, 0) = reportID And EditorHasPending() Then Err.Raise vbObjectError + 9621, "modSignoff.DisposeReportDrafts", "Save or discard pending edits first."
    End If
    signoff = ReportCurrentSignoff(reportID)
    ids = DisposableRevisions(reportID)
    operation = BeginOperation("Dispose report drafts", "Report " & reportID & "; signed record " & signoff)
    If ids <> "" Then
        revisions = UBound(Split(ids, ",")) + 1
        db.Execute "DELETE FROM RevisionValue WHERE RevisionID IN (" & ids & ")", dbFailOnError: values = db.RecordsAffected
        db.Execute "DELETE FROM RevisionDate WHERE RevisionID IN (" & ids & ")", dbFailOnError: values = values + db.RecordsAffected
        db.Execute "DELETE FROM RevisionChoice WHERE RevisionID IN (" & ids & ")", dbFailOnError: values = values + db.RecordsAffected
        db.Execute "DELETE FROM RevisionChange WHERE RevisionID IN (" & ids & ")", dbFailOnError: changes = db.RecordsAffected
        db.Execute "UPDATE ReportRevision SET ValuesDisposed=True WHERE RevisionID IN (" & ids & ")", dbFailOnError
    End If
    'A kept revision's change list holds the previous draft's values when its parent
    'is a disposed draft; a change between two signed revisions stays.
    db.Execute "DELETE FROM RevisionChange WHERE RevisionID IN (SELECT RevisionID FROM ReportRevision WHERE ReportID=" & reportID & " AND Not Nz(ValuesDisposed,False) AND ParentRevisionID IN (SELECT RevisionID FROM ReportRevision WHERE ReportID=" & reportID & " AND Nz(ValuesDisposed,False)))", dbFailOnError
    changes = changes + db.RecordsAffected
    db.Execute "DELETE FROM RecoveryDraft WHERE ReportID=" & reportID, dbFailOnError: recovery = db.RecordsAffected
    db.Execute "DELETE FROM ReportNoteDraft WHERE ReportID=" & reportID, dbFailOnError: notes = db.RecordsAffected
    db.Execute "DELETE FROM ImportRawValue WHERE ImportReceiptID IN (SELECT ImportReceiptID FROM ImportReceipt WHERE ReportID=" & reportID & ")", dbFailOnError: raw = db.RecordsAffected
    Set d = db.OpenRecordset("DraftDisposal", dbOpenDynaset)
    d.AddNew: d!DisposalGuid = NewIdentity(): d!ReportID = reportID: d!SignoffID = signoff: d!DisposedBy = MajorAuthor(): d!DisposedAt = Now()
    d!RevisionsCleared = revisions: d!ValueRowsRemoved = values: d!ChangeRowsRemoved = changes: d!RecoveryRemoved = recovery: d!NoteDraftsRemoved = notes: d!RawValuesRemoved = raw
    d.Update: d.Bookmark = d.LastModified: DisposeReportDrafts = d!DisposalID: d.Close
    ReportEventRecord reportID, "Drafts disposed", operation
    FinishOperation operation, "Completed", revisions & " draft revisions emptied; signed revisions, conversations, reviews and accounting kept. Existing backups and exports still hold the drafts."
    w.CommitTrans: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modSignoff", "DisposeReportDrafts", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    SwiftRollback w, swiftPrimary
    If Not d Is Nothing Then SwiftClose d, swiftPrimary, "d"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function
Public Function RevisionValuesDisposed(ByVal revisionID As Long) As Boolean
    RevisionValuesDisposed = Nz(MajorScalar("SELECT ValuesDisposed FROM ReportRevision WHERE RevisionID=" & revisionID), False)
End Function

'Tests answer the reopen prompt here instead of a message box (one use).
Public Sub UXSignedTestAnswer(ByVal decision As Long, Optional ByVal reason As String = "")
    testDecision = decision: testReason = reason: testUsed = False
End Sub
'True when editing may continue. A signed report asks whether to reopen it for
'correction; No returns False so the caller leaves the field unchanged.
Public Function UXSignedEditGuard() As Boolean
    Dim id As Long, decision As Long, reason As String
    UXSignedEditGuard = True
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then Exit Function
    id = Nz(Forms!frmEditor!txtReportID, 0)
    If id = 0 Then Exit Function
    If Not ReportIsSigned(id) Then Exit Function
    If testDecision <> 0 And Not testUsed Then
        decision = testDecision: reason = testReason: testUsed = True: testDecision = 0
    Else
        decision = MsgBox(Replace(SignedReportCaption(id), " Editing a field asks whether to reopen it for correction.", "") & vbCrLf & vbCrLf & "Reopen it for correction? No leaves the field unchanged.", vbYesNo + vbQuestion + vbDefaultButton2, "Reopen signed report")
        If decision = vbYes Then reason = InputBox("Reason for the correction (optional):", "Reopen for correction")
    End If
    If decision <> vbYes Then UXSignedEditGuard = False: Exit Function
    On Error GoTo Failed
    Call ReopenSignedReport(id, reason)
    Forms!frmEditor!cmdSave.Enabled = True: Forms!frmEditor!cboKind.Locked = False
    Call UXHeaderReviewState
    Forms!frmEditor!uxLocked.Caption = "Reopened for correction. Mark the corrected report signed again when it is re-signed."
    Exit Function
Failed:
    Forms!frmEditor!uxError.Caption = "Could not reopen: " & Err.Description
    UXSignedEditGuard = False
End Function

'--- Mark signed dialog -------------------------------------------------------
Public Function UXSignoffOpen(Optional ByVal reportID As Long = 0) As Boolean
    Dim f As Access.Form
    On Error GoTo Failed
    If reportID = 0 And CurrentProject.AllForms("frmEditor").IsLoaded Then
        reportID = Nz(Forms!frmEditor!txtReportID, 0)
        If reportID = 0 Then Err.Raise vbObjectError + 9630, "modSignoff.UXSignoffOpen", "Save the report before marking it signed."
        If EditorHasPending() Then Err.Raise vbObjectError + 9631, "modSignoff.UXSignoffOpen", "Save or discard pending edits first. Mark signed records the saved version."
    End If
    If reportID = 0 Then Err.Raise vbObjectError + 9630, "modSignoff.UXSignoffOpen", "Select a saved report."
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Nz(Forms!frmEditor!txtReportID, 0) = reportID And EditorHasPending() Then Err.Raise vbObjectError + 9631, "modSignoff.UXSignoffOpen", "Save or discard pending edits first. Mark signed records the saved version."
    End If
    If ReportIsSigned(reportID) Then Err.Raise vbObjectError + 9602, "modSignoff.UXSignoffOpen", "This report is already marked signed."
    signoffReport = reportID: signoffRevision = CurrentRevision(reportID): signoffGroup = 0
    DoCmd.OpenForm "frmUXSignoff": Set f = Forms!frmUXSignoff
    f!lblReport.Caption = UXMemberLabel(reportID) & " | " & Nz(MajorScalar("SELECT FormKind FROM Report WHERE ReportID=" & reportID), "") & " | saved revision " & Nz(MajorScalar("SELECT RevisionNumber FROM ReportRevision WHERE RevisionID=" & signoffRevision), "")
    f!txtSignedOn = Date: f!chkDispose = False
    f!lblStatus.Caption = "Records that this saved version was signed outside Swift. It is not an official signature or submission. The report becomes read only; editing a field later asks whether to reopen it."
    UXSignoffOpen = True: Exit Function
Failed:
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Forms!frmEditor!uxError.Caption = Err.Description Else MsgBox Err.Description, vbExclamation, "Mark signed"
End Function
Public Function UXSignoffSave(Optional ByVal confirmedDisposal As Boolean = False) As Boolean
    Dim f As Access.Form, dispose As Boolean
    On Error GoTo Failed
    Set f = Forms!frmUXSignoff
    If Not IsDate(f!txtSignedOn) Then Err.Raise vbObjectError + 9632, "modSignoff.UXSignoffSave", "Enter the date the report was signed."
    dispose = Nz(f!chkDispose, False)
    If dispose And Not confirmedDisposal Then
        If MsgBox("Dispose of this report's drafts now? Superseded draft revisions, recovery checkpoints, unposted note drafts and imported source values are removed. The signed version, conversations, reviews and accounting stay. This cannot be undone, and existing backups and exports still hold the drafts.", vbYesNo + vbExclamation + vbDefaultButton2, "Dispose of drafts") <> vbYes Then dispose = False
    End If
    If signoffGroup > 0 Then
        UXSignoffSave = UXSignoffSaveGroup(CDate(f!txtSignedOn), dispose): Exit Function
    End If
    Call MarkReportSigned(signoffReport, signoffRevision, CDate(f!txtSignedOn))
    If dispose Then Call DisposeReportDrafts(signoffReport)
    DoCmd.Close acForm, f.Name
    UXSignoffRefresh signoffReport
    UXSignoffSave = True: Exit Function
Failed:
    f!lblStatus.Caption = Err.Description
End Function
Private Function UXSignoffSaveGroup(ByVal signedOn As Date, ByVal dispose As Boolean) As Boolean
    Dim r As DAO.Recordset, before As String, marked As Long, id As Variant, ids As New Collection
    Set r = CurrentDb.OpenRecordset("SELECT r.ReportID FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & signoffGroup & " AND r.IsArchived=False AND r.WorkflowState In ('Draft','Approved')", dbOpenSnapshot)
    Do While Not r.EOF: ids.Add CLng(r!ReportID): r.MoveNext: Loop
    r.Close
    If signoffCompleting Then
        SetGroupWorkflow signoffGroup, signoffStamp, "Completed", signedOn
    Else
        marked = MarkGroupReportsSigned(signoffGroup, signedOn)
    End If
    If dispose Then
        For Each id In ids: Call DisposeReportDrafts(CLng(id)): Next
    End If
    DoCmd.Close acForm, "frmUXSignoff"
    If CurrentProject.AllForms("frmUXGroups").IsLoaded Then UXGroupList
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then UXGridLoad
    UXSignoffSaveGroup = True
End Function
Public Function UXDisposeDrafts(Optional ByVal reportID As Long = 0, Optional ByVal confirmed As Boolean = False) As Boolean
    On Error GoTo Failed
    If reportID = 0 And CurrentProject.AllForms("frmEditor").IsLoaded Then reportID = Nz(Forms!frmEditor!txtReportID, 0)
    If DisposableRevisions(reportID) = "" And Nz(MajorScalar("SELECT Count(*) FROM RecoveryDraft WHERE ReportID=" & reportID), 0) = 0 Then Err.Raise vbObjectError + 9622, "modSignoff.UXDisposeDrafts", "This report has no drafts left to dispose of."
    If Not confirmed Then
        If MsgBox("Dispose of this signed report's drafts? Superseded draft revisions, recovery checkpoints, unposted note drafts and imported source values are removed. The signed version, conversations, reviews and accounting stay. This cannot be undone.", vbYesNo + vbExclamation + vbDefaultButton2, "Dispose of drafts") <> vbYes Then Exit Function
    End If
    Call DisposeReportDrafts(reportID)
    UXSignoffRefresh reportID
    UXDisposeDrafts = True: Exit Function
Failed:
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Forms!frmEditor!uxError.Caption = Err.Description Else MsgBox Err.Description, vbExclamation, "Dispose of drafts"
End Function
'The change is already committed, so a refresh failure is reported, not raised.
Private Sub UXSignoffRefresh(ByVal reportID As Long)
    Dim stage As String, failure As Object
    On Error GoTo Failed
    stage = "Editor refresh"
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Nz(Forms!frmEditor!txtReportID, 0) = reportID Then
            'Fields stay open when signed (the first edit asks), so only the actions change.
            Forms!frmEditor!cmdSave.Enabled = Not ReportIsSigned(reportID): Forms!frmEditor!cboKind.Locked = ReportIsSigned(reportID)
            If Not UXHeaderReviewState() Then
                Set failure = SwiftLastFailureContext()
                failure("Transaction") = "Committed"
                GoTo RefreshFailed
            End If
            Forms!frmEditor!uxLocked.Caption = SignedReportCaption(reportID)
        End If
    End If
    stage = "Comparison refresh"
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then UXGridLoad
    Exit Sub
Failed:
    Set failure = SwiftFailure("modSignoff", "UXSignoffRefresh", stage, Err.Number, Err.Description, Err.Source, "Committed")
RefreshFailed:
    On Error Resume Next
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        Forms!frmEditor!uxError.Caption = Left$("Saved; " & LCase$(stage) & " failed. Reopen the report. " & SwiftRedact(CStr(failure("Description"))), 255)
        Forms!frmEditor!uxError.ControlTipText = Left$(SwiftFailureText(failure), 255)
    ElseIf CurrentProject.AllForms("frmUXComparison").IsLoaded Then
        Forms!frmUXComparison!lblStatus.Caption = Left$("Saved; " & SwiftFailureText(failure), 255)
    End If
    If stage = "Editor refresh" And CurrentProject.AllForms("frmUXComparison").IsLoaded Then UXGridLoad
End Sub

'--- Summary groups ------------------------------------------------------------
'Mark every active draft or approved report in a group signed on one date.
'Already signed reports are left as they are. Returns the number marked.
Public Function MarkGroupReportsSigned(ByVal groupID As Long, ByVal signedOn As Date) As Long
    Dim r As DAO.Recordset, w As DAO.Workspace
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If DCount("*", "GroupMember", "GroupID=" & groupID & " AND ReportID=" & Nz(Forms!frmEditor!txtReportID, 0)) > 0 And EditorHasPending() Then Err.Raise vbObjectError + 9640, "modSignoff.MarkGroupReportsSigned", "Save or discard pending edits in the open report first."
    End If
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT r.ReportID,r.CurrentRevisionID FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & groupID & " AND r.IsArchived=False AND r.WorkflowState In ('Draft','Approved') ORDER BY r.ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        Call MarkReportSigned(r!ReportID, r!CurrentRevisionID, signedOn)
        MarkGroupReportsSigned = MarkGroupReportsSigned + 1: r.MoveNext
    Loop
    r.Close
    w.CommitTrans: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modSignoff", "MarkGroupReportsSigned", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function
Public Function ReopenGroupReports(ByVal groupID As Long, ByVal reason As String) As Long
    Dim r As DAO.Recordset, ids As New Collection, id As Variant
    Set r = CurrentDb.OpenRecordset("SELECT r.ReportID FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & groupID & " AND r.WorkflowState='Signed'", dbOpenSnapshot)
    Do While Not r.EOF: ids.Add CLng(r!ReportID): r.MoveNext: Loop
    r.Close
    For Each id In ids: Call ReopenSignedReport(CLng(id), reason): Next
    ReopenGroupReports = ids.Count
End Function
Public Function GroupSignedCount(ByVal groupID As Long) As Long
    GroupSignedCount = Nz(MajorScalar("SELECT Count(*) FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & groupID & " AND r.WorkflowState='Signed'"), 0)
End Function
Public Function GroupSignableCount(ByVal groupID As Long) As Long
    GroupSignableCount = Nz(MajorScalar("SELECT Count(*) FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & groupID & " AND r.IsArchived=False AND r.WorkflowState In ('Draft','Approved')"), 0)
End Function
'Open the Mark signed dialog for a whole group. completing: the dialog also sets the
'group to Completed (stamp is the group list's state stamp).
Public Function UXSignoffOpenGroup(ByVal groupID As Long, ByVal completing As Boolean, Optional ByVal stamp As String = "") As Boolean
    Dim f As Access.Form
    On Error GoTo Failed
    If groupID = 0 Then Err.Raise vbObjectError + 9641, "modSignoff.UXSignoffOpenGroup", "Select a summary group."
    If Not completing And GroupSignableCount(groupID) = 0 Then Err.Raise vbObjectError + 9642, "modSignoff.UXSignoffOpenGroup", "Every report in this group is already signed or archived."
    signoffReport = 0: signoffRevision = 0: signoffGroup = groupID: signoffCompleting = completing: signoffStamp = stamp
    DoCmd.OpenForm "frmUXSignoff": Set f = Forms!frmUXSignoff
    f!lblReport.Caption = Nz(DLookup("Title", "SummaryGroup", "GroupID=" & groupID), "Summary group") & " | " & GroupSignableCount(groupID) & " reports to mark signed"
    f!txtSignedOn = Date: f!chkDispose = False
    f!lblStatus.Caption = IIf(completing, "Completing this group marks all of its draft and approved reports signed on this date. Reopening the group later reopens them.", "Marks all of this group's draft and approved reports signed on this date. Already signed reports are unchanged.")
    UXSignoffOpenGroup = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Mark signed"
End Function