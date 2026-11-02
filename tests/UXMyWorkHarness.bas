Attribute VB_Name = "UXMyWorkHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 5960, , message
    checks = checks + 1
End Sub
Private Function RetiredResolutionRejected(ByVal id As Long) As Boolean
    On Error GoTo Rejected
    SetReviewThreadState id, True
    Exit Function
Rejected:
    RetiredResolutionRejected = True
End Function
Private Sub CompareRows(ByVal pid As Long, ByVal history As Boolean)
    Dim before As DAO.Recordset, after As DAO.Recordset, i As Long, ignored As Long, visible As Object, rowCount As Long
    Call UXPrepareWorkRows(pid, history)
    ignored = UXWorkCount()
    If Not history Then Set visible = UXWorkVisibleRows(pid)
    Set before = CurrentDb.OpenRecordset(UXWorkSQL(pid, history), dbOpenSnapshot)
    Set after = CurrentDb.OpenRecordset(UXWorkSQL(pid, history, True), dbOpenSnapshot)
    Do While Not before.EOF
        rowCount = rowCount + 1
        If Not history Then Check visible.Exists(CStr(before!WorkKey)), "Batched inclusion omitted key"
        Check Not after.EOF, "Prepared queue omitted a row"
        For i = 0 To before.Fields.Count - 1
            Check SameFieldValue(before.Fields(i).Value, after.Fields(i).Value), "Queue cell/order differs: " & before.Fields(i).Name
        Next
        before.MoveNext: after.MoveNext
    Loop
    If Not history Then Check visible.Count = rowCount, "Batched inclusion added keys"
    Check after.EOF, "Prepared queue added rows"
    before.Close: after.Close
End Sub
Private Sub CheckCount()
    Dim r As DAO.Recordset, profiles As DAO.Recordset, expected As Long
    Set r = CurrentDb.OpenRecordset(UXWorkSQL(MajorAuthor()), dbOpenSnapshot)
    If Not r.EOF Then r.MoveLast: expected = r.RecordCount
    r.Close
    Check UXWorkCount() = expected, "Overview count differs from visible queue"
    Set profiles = CurrentDb.OpenRecordset("SELECT UserProfileID FROM UserProfile", dbOpenSnapshot)
    Do While Not profiles.EOF
        CompareRows profiles!UserProfileID, False
        CompareRows profiles!UserProfileID, True
        profiles.MoveNext
    Loop
    profiles.Close
    If CurrentProject.AllForms("frmMyWork").IsLoaded Then
        UXPrepareWorkRows MajorAuthor(), Nz(Forms!frmMyWork!chkAll, False)
    Else
        UXPrepareWorkRows MajorAuthor(), False
    End If
End Sub
Public Function TestUXMyWork() As String
    Dim author As Long, other As Long, id As Long, task As Long, template As Long, run As Long, thread As Long, reply As Long, oldRevision As Long, f As Access.Form, width As Variant, extra As Long, group As Long, unrelated As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC INBOX", "Author"
    author = MajorAuthor(): other = AddContributor("Second reviewer")
    unrelated = AddContributor("Unrelated contributor")
    id = CommitDraft(0, "FITREP", "Inbox fixture", "Synthetic tester", #1/1/2026#, #9/30/2026#, "", "Test", "")
    oldRevision = CurrentRevision(id)
    task = AssignReview(id, oldRevision, author, #1/1/2026#, "Check all saved blocks")
    template = CreateRouteTemplate("Ordered fixture", author & "," & other, "Read the saved revision")
    run = StartRoute(id, template)
    thread = AddReviewComment(id, oldRevision, "Unresolved synthetic feedback")
    reply = AddReviewComment(id, oldRevision, "Reply stays in this thread", thread)
    Check UXWorkVisible("Assignment", task, author), "Assignment missing"
    Check UXWorkVisible("Route", run, author), "Current step missing"
    Check Not UXWorkVisible("Route", run, other), "Future reviewer sees an actionable step"
    Check Not UXWorkVisible("Comment", thread, author), "Notes must not become approval tasks"
    Check Not UXWorkVisible("Comment", thread, other), "Unrelated future reviewer thread included"
    Call CheckCount
    Check UXWorkCount() = 2, "Replies duplicated or inbox item missing"
    SwitchContributor other
    reply = AddReviewComment(id, oldRevision, "Participation through reply only", thread)
    SwitchContributor author
    Check Not UXWorkVisible("Comment", thread, other), "Replies must stay out of approval queue"
    Check Not UXWorkVisible("Comment", thread, unrelated), "Unrelated contributor included"
    Call CheckCount
    DoCmd.OpenForm "frmMyWork": Set f = Forms!frmMyWork
    Check f!lstWork.ListCount = 3, "Unified query/headings failed"
    f!lstWork = "Assignment:" & task: Call UXWorkSelection
    Check f!cmdAccept.Enabled And f!cmdReturn.Enabled And f!cmdCancel.Enabled, "Current assignment actions"
    Check InStr(f!txtHistory, "Saved revision 1") > 0 And InStr(f!txtHistory, "revision ID") = 0, "Raw revision ID exposed"
    Call CommitDraft(id, "FITREP", "Inbox fixture", "Synthetic tester", #1/1/2026#, #9/30/2026#, "Changed", "Test", "")
    Call CheckCount
    Call UXLoadWorkQueue
    Check f!lstWork = "Assignment:" & task, "Queue selection lost"
    ' F01: the holder's own save keeps both reviews current.
    Check f!cmdAccept.Enabled And f!cmdReturn.Enabled And f!cmdCancel.Enabled, "Holder save disabled the assignment decisions"
    Check InStr(UXWorkText("Route", run, "Action"), "forward or send back") > 0 And RouteIsCurrent(run), "Holder save staled the route"
    Check InStr(UXWorkText("Comment", thread, "Status"), "earlier revision") > 0, "Earlier thread misrepresented"
    MigrateReportMessages id
    Check UXOpenReportTool(id, "Notes"), "Notes navigation failed"
    Check UXToolsForm()!lstMessages.ListCount > 1, "Historical conversation missing"
    Check UXWorkBack(), "Notes return failed"
    Call DecideRouteStep(run, "Accepted", "First step complete")
    Check Not UXWorkVisible("Route", run, author) And UXWorkVisible("Route", run, other), "Ordered step did not advance"
    Call SwitchContributor(other): Call LoadMyWork
    f!lstWork = "Route:" & run: Call UXWorkSelection
    Check Not f!cmdAccept.Enabled And UXWorkOpen(), "Route should use focused decisions"
    Check UXReviewForm()!cboRun = run And UXReviewForm()!cmdAccept.Enabled, "Wrong route or reviewer opened"
    Check UXWorkBack(), "Review return failed"
    Call DecideRouteStep(run, "Accepted", "Final step complete")
    Check UXWorkVisible("Route", run, other) And InStr(UXWorkText("Route", run, "Action"), "local approval") > 0, "Local approval work missing"
    Call ApproveSavedRoute(run, "Synthetic approval")
    Check Not UXWorkVisible("Route", run, other), "Approved route remains actionable"
    Check Not UXWorkVisible("Comment", thread, other), "Approved report notes became approval tasks"
    Check RetiredResolutionRejected(thread), "Retired resolution API accepted a write"
    Call CheckCount
    Call LoadMyWork
    Check f!lstWork.ListCount <= 1 And InStr(f!uxWorkStatus.Caption, "No work") > 0, "Empty inbox guidance"
    f!chkAll = True: Call LoadMyWork
    Check f!lstWork.ListCount >= 3, "Completed history unavailable"
    f!lstWork = "Assignment:" & task: Call UXWorkSelection
    Check Not f!cmdAccept.Enabled, "Local approval assignment decision enabled"
    For Each width In Array(9000, 14000, 20000)
        Check UXResizeWorkspace("frmMyWork", CLng(width), 10000), "Inbox resize failed"
        Check f!lstWork.Left + f!lstWork.Width < CLng(width), "Inbox horizontal bounds"
        Check ColumnTotal(f!lstWork.ColumnWidths) <= f!lstWork.Width, "Inbox columns wider than the list at " & width & ": " & f!lstWork.ColumnWidths & " in " & f!lstWork.Width
    Next
    Call ArchiveDraft(id, True)
    Call CheckCount
    extra = CommitDraft(0, "EVAL", "Returned fixture", "Return tester", Null, Null, "")
    task = AssignReview(extra, CurrentRevision(extra), MajorAuthor(), Null, "Return this")
    FinishReview task, "Returned", "Needs revision"
    Call CheckCount
    task = AssignReview(extra, CurrentRevision(extra), MajorAuthor(), Null, "Cancel this")
    FinishReview task, "Cancelled", "Cancelled fixture"
    Call CheckCount
    group = CreateSummaryGroup("Work row group", "EVAL", "E5")
    AddGroupMember group, extra, "E5"
    template = CreateRouteTemplate("Group route fixture", CStr(MajorAuthor()), "Review group")
    run = StartRoute(extra, template, group)
    Call CheckCount
    id = CommitDraft(0, "EVAL", "Membership change", "Another tester", Null, Null, "")
    AddGroupMember group, id, "E5"
    Call CheckCount
    Check Not RouteIsCurrent(run), "Group change did not stale route"
    ArchiveDraft extra, True
    Call CheckCount
    TestUXMyWork = "PASS: " & checks & " unified inbox scope, ordered steps, comments, staleness, navigation and layout assertions"
    Exit Function
Failed:
    TestUXMyWork = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
Private Function ColumnTotal(ByVal widths As String) As Long
    Dim part As Variant
    For Each part In Split(widths, ";"): ColumnTotal = ColumnTotal + Val(part): Next
End Function
