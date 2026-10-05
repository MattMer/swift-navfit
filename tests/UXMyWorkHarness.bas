Attribute VB_Name = "UXMyWorkHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 5960, , message
    checks = checks + 1
End Sub
Public Function TestUXMyWork() As String
    Dim author As Long, other As Long, id As Long, task As Long, template As Long, run As Long, thread As Long, reply As Long, oldRevision As Long, f As Access.Form, width As Variant
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC INBOX", "Author"
    author = MajorAuthor(): other = AddContributor("Second reviewer")
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
    Check UXWorkVisible("Comment", thread, author), "Author thread missing"
    Check Not UXWorkVisible("Comment", thread, other), "Unrelated future reviewer thread included"
    Check UXWorkCount() = 3, "Replies duplicated or inbox item missing"
    DoCmd.OpenForm "frmMyWork": Set f = Forms!frmMyWork
    Check f!lstWork.ListCount = 4, "Unified query/headings failed"
    f!lstWork = "Assignment:" & task: Call UXWorkSelection
    Check f!cmdAccept.Enabled And f!cmdReturn.Enabled And f!cmdCancel.Enabled, "Current assignment actions"
    Check InStr(f!txtHistory, "Saved revision 1") > 0 And InStr(f!txtHistory, "revision ID") = 0, "Raw revision ID exposed"
    Call CommitDraft(id, "FITREP", "Inbox fixture", "Synthetic tester", #1/1/2026#, #9/30/2026#, "Changed", "Test", "")
    Call UXLoadWorkQueue
    Check f!lstWork = "Assignment:" & task, "Queue selection lost"
    Check Not f!cmdAccept.Enabled And Not f!cmdReturn.Enabled And f!cmdCancel.Enabled, "Stale decisions enabled"
    Check InStr(UXWorkText("Route", run, "Action"), "restart") > 0, "Stale route next action"
    Check InStr(UXWorkText("Comment", thread, "Status"), "earlier revision") > 0, "Earlier thread misrepresented"
    f!lstWork = "Comment:" & thread: Call UXWorkSelection
    Check UXWorkOpen(), "Comment navigation failed"
    Check Forms!frmReviewComments!lstThreads = thread, "Wrong comment selected"
    Check InStr(Forms!frmReviewComments!txtHistory, "saved revision 1") > 0, "Historical baseline lost"
    Call SetReviewThreadState(thread, True)
    Check Not UXThreadOpen(thread) And Not UXWorkVisible("Comment", thread, author), "Resolved thread remains pending"
    Call SetReviewThreadState(thread, False)
    Check UXThreadOpen(thread) And UXWorkVisible("Comment", thread, author), "Reopened thread missing"
    DoCmd.Close acForm, "frmReviewComments", acSaveNo
    run = RestartStaleRoute(run, CurrentRevision(id), "")
    Call DecideRouteStep(run, "Accepted", "First step complete")
    Check Not UXWorkVisible("Route", run, author) And UXWorkVisible("Route", run, other), "Ordered step did not advance"
    Call SwitchContributor(other): Call LoadMyWork
    f!lstWork = "Route:" & run: Call UXWorkSelection
    Check Not f!cmdAccept.Enabled And UXWorkOpen(), "Route should use focused decisions"
    Check Forms!frmUXReview!cboRun = run And Forms!frmUXReview!cmdAccept.Enabled, "Wrong route or reviewer opened"
    DoCmd.Close acForm, "frmUXReview", acSaveNo
    Call DecideRouteStep(run, "Accepted", "Final step complete")
    Check UXWorkVisible("Route", run, other) And InStr(UXWorkText("Route", run, "Action"), "local approval") > 0, "Local approval work missing"
    Call ApproveSavedRoute(run, "Synthetic approval")
    Check Not UXWorkVisible("Route", run, other), "Approved route remains actionable"
    Call LoadMyWork
    Check f!lstWork.ListCount <= 1 And InStr(f!uxWorkStatus.Caption, "No work") > 0, "Empty inbox guidance"
    f!chkAll = True: Call LoadMyWork
    Check f!lstWork.ListCount >= 4, "Completed history unavailable"
    f!lstWork = "Assignment:" & task: Call UXWorkSelection
    Check Not f!cmdAccept.Enabled, "Local approval assignment decision enabled"
    For Each width In Array(9000, 14000, 20000)
        Check UXWorkLayout(CLng(width), 10000), "Inbox resize failed"
        Check f!lstWork.Left + f!lstWork.Width < CLng(width), "Inbox horizontal bounds"
    Next
    TestUXMyWork = "PASS: " & checks & " unified inbox scope, ordered steps, comments, staleness, navigation and layout assertions"
    Exit Function
Failed:
    TestUXMyWork = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
