Attribute VB_Name = "ReviewCommentsHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 1750, , message
    checks = checks + 1
End Sub
Private Function Rejected(ByVal reportID As Long, ByVal revisionID As Long, ByVal body As String, Optional ByVal threadID As Long = 0) As Boolean
    On Error GoTo Failed
    Call AddReviewComment(reportID, revisionID, body, threadID): Exit Function
Failed:
    Rejected = True
End Function
Public Function TestReviewComments() As String
    Dim reportID As Long, other As Long, revisionID As Long, root As Long, reply As Long, newer As Long
    Dim before As Long, body As String, r As DAO.Recordset, pid As Long
    On Error GoTo Failed
    Call ConfigureWorkspace("SYNTHETIC COMMENTS", "First Reviewer")
    reportID = CommitDraft(0, "EVAL", "Test", "Member", Null, Null, "Notes", "Initial", "Official narrative")
    other = CommitDraft(0, "FITREP", "Other", "Other", Null, Null, "", "Initial", "Other narrative")
    revisionID = DLookup("CurrentRevisionID", "Report", "ReportID=" & reportID)
    before = DCount("*", "ReportRevision")
    body = String$(10000, "Q") & vbCrLf & "Reviewer's exact text"
    root = AddReviewComment(reportID, revisionID, body)
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ReviewComment WHERE CommentID=" & root, dbOpenSnapshot)
    Check StrComp(r!CommentBody, body, vbBinaryCompare) = 0, "Long comment changed"
    Check r!RevisionID = revisionID And IsNull(r!ThreadID), "Thread revision missing"
    r.Close
    Check Rejected(reportID, revisionID, vbCrLf & vbTab & " "), "Blank comment accepted"
    Check Rejected(other, revisionID, "Cross report"), "Cross-report revision accepted"
    Check Rejected(other, DLookup("CurrentRevisionID", "Report", "ReportID=" & other), "Cross thread", root), "Cross-report reply accepted"
    Set r = CurrentDb.OpenRecordset("UserProfile", dbOpenDynaset)
    r.AddNew: r!ProfileGuid = NewIdentity(): r!DisplayName = "Second Reviewer": r.Update: r.Bookmark = r.LastModified: pid = r!UserProfileID: r.Close
    CurrentDb.Execute "UPDATE Workspace SET ActiveProfileID=" & pid, dbFailOnError
    reply = AddReviewComment(reportID, revisionID, "Reply", root)
    Check DLookup("AuthorProfileID", "ReviewComment", "CommentID=" & reply) = pid, "Reply contributor wrong"
    Check SetReviewThreadState(root, True), "Resolve failed"
    Check Not SetReviewThreadState(root, True), "Repeated resolve appended event"
    Check Rejected(reportID, revisionID, "Reply to resolved", root), "Resolved reply accepted"
    Check SetReviewThreadState(root, False), "Reopen failed"
    Check AddReviewComment(reportID, revisionID, "After reopen", root) > 0, "Reopened reply failed"
    Check DCount("*", "ReportRevision") = before, "Comments changed report revisions"
    Set r = CurrentDb.OpenRecordset("SELECT OfficialComments FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    Check r!OfficialComments = "Official narrative", "Comment entered official narrative": r.Close
    DoCmd.OpenForm "frmEditor", , , , , , CStr(reportID)
    Forms!frmEditor!txtTitle = "Unsaved title"
    Check CommentsFromEditor(), "Editor action failed"
    Check Forms!frmReviewComments!lstThreads.ListCount = 1 + IIf(Forms!frmReviewComments!lstThreads.ColumnHeads, 1, 0), "Thread list count wrong"
    Forms!frmReviewComments!lstThreads = root
    Check ShowReviewThread(), "Thread display failed"
    Check InStr(Forms!frmReviewComments!txtHistory, body) > 0 And InStr(Forms!frmReviewComments!txtHistory, "Resolved") > 0 And InStr(Forms!frmReviewComments!txtHistory, "Reopened") > 0, "History missing"
    Check Forms!frmReviewComments!txtHistory.Locked, "History editable"
    Check ReturnFromReviewComments(), "Return failed"
    Check Forms!frmEditor!txtTitle = "Unsaved title", "Unsaved editor altered"
    Forms!frmReviewComments!txtCompose = "UI reply": Forms!frmReviewComments!lstThreads = root
    Check PostReviewUI(True), "UI reply failed"
    Check IsNull(Forms!frmReviewComments!txtCompose), "Compose not cleared after post"
    Call CommitDraft(reportID, "EVAL", "New saved", "Member", Null, Null, "Notes", "Later", "New official")
    newer = DLookup("CurrentRevisionID", "Report", "ReportID=" & reportID)
    Check Rejected(reportID, newer, "Wrong revision reply", root), "Reply changed thread revision"
    Forms!frmReviewComments!txtCompose = "New thread on latest saved"
    Check PostReviewUI(), "UI new thread failed"
    Check DCount("*", "ReviewComment", "ReportID=" & reportID & " AND RevisionID=" & newer & " AND ThreadID Is Null") = 1, "New thread did not use latest revision"
    Call FixtureCloseEditor()
    Call ArchiveDraft(reportID, True)
    Check Rejected(reportID, newer, "Archived comment"), "Archived comment accepted"
    Check OpenReviewComments(reportID), "Archived history inaccessible"
    Forms!frmReviewComments!lstThreads = root: Check ShowReviewThread(), "Archived history missing"
    DoCmd.Close acForm, "frmReviewComments", acSaveNo
    TestReviewComments = "PASS: " & checks & " review comment assertions": Exit Function
Failed:
    TestReviewComments = "FAIL: " & Err.Number & " " & Err.Description
End Function

Private Sub FixtureCloseEditor()
    ' Preserve the old abrupt-close semantics for persistence/recovery fixtures.
    ' Real save/discard/cancel exits are exercised by UXWorkflowHarness.
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Sub
