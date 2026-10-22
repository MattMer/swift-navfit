Attribute VB_Name = "modReviewComments"
Option Compare Database
Option Explicit
Private composingReply As Boolean
Private Function CommentProfile() As Long
    CommentProfile = ActiveContributorID()
End Function
Private Sub CommentTarget(ByVal reportID As Long, ByVal revisionID As Long)
    If DCount("*", "ReportRevision", "ReportID=" & reportID) = 0 Then Err.Raise vbObjectError + 1701, , "Choose a saved revision belonging to this report."
    If Nz(DLookup("IsArchived", "Report", "ReportID=" & reportID), True) Then Err.Raise vbObjectError + 1702, , "Recover archived reports before adding review activity."
End Sub
Private Function ThreadResolved(ByVal threadID As Long) As Boolean
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT TOP 1 EventKind FROM ReviewCommentEvent WHERE ThreadID=" & threadID & " ORDER BY CommentEventID DESC", dbOpenSnapshot)
    If Not r.EOF Then ThreadResolved = (r!EventKind = "Resolved")
    r.Close
End Function
Public Function AddReviewComment(ByVal reportID As Long, ByVal revisionID As Long, ByVal body As String, Optional ByVal threadID As Long = 0) As Long
    Dim r As DAO.Recordset, pid As Long
    pid = CommentProfile(): Call CommentTarget(reportID, revisionID)
    If Len(Trim$(Replace(Replace(Replace(body, vbCr, ""), vbLf, ""), vbTab, ""))) = 0 Then Err.Raise vbObjectError + 1703, , "Enter a review comment."
    If threadID <> 0 Then
        If DCount("*", "ReviewComment", "CommentID=" & threadID & " AND ThreadID Is Null AND ReportID=" & reportID) = 0 Then Err.Raise vbObjectError + 1704, , "Reply must belong to the selected report and revision thread."
    End If
    Set r = CurrentDb.OpenRecordset("ReviewComment", dbOpenDynaset)
    r.AddNew: r!CommentGuid = NewIdentity(): r!ReportID = reportID: r!RevisionID = revisionID
    If threadID > 0 Then r!ThreadID = threadID
    r!AuthorProfileID = pid: r!CommentAt = Now(): r!CommentBody = body
    r.Update: r.Bookmark = r.LastModified: AddReviewComment = r!CommentID: r.Close
End Function
Public Function SetReviewThreadState(ByVal threadID As Long, ByVal resolved As Boolean) As Boolean
    Dim r As DAO.Recordset, reportID As Long, revisionID As Long, pid As Long
    pid = CommentProfile()
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ReviewComment WHERE CommentID=" & threadID & " AND ThreadID Is Null", dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 1706, , "Select a review thread."
    reportID = r!ReportID: revisionID = r!RevisionID: r.Close
    Call CommentTarget(reportID, revisionID)
    If ThreadResolved(threadID) = resolved Then Exit Function
    Set r = CurrentDb.OpenRecordset("ReviewCommentEvent", dbOpenDynaset)
    r.AddNew: r!ThreadID = threadID: r!AuthorProfileID = pid: r!EventAt = Now()
    If resolved Then r!EventKind = "Resolved" Else r!EventKind = "Reopened"
    r.Update: r.Close: SetReviewThreadState = True
    Call UXWorkChanged
End Function
Public Function OpenReviewComments(ByVal reportID As Long) As Boolean
 OpenReviewComments = UXOpenReportTool(reportID, "Notes")
End Function
Public Function CommentsFromEditor() As Boolean
 CommentsFromEditor = UXToolsChoose("Notes")
End Function
Public Function LoadReviewThreads() As Boolean
 LoadReviewThreads = UXMessageRefresh()
End Function
Public Function ShowReviewThread() As Boolean
 ShowReviewThread = UXMessageRead()
End Function
Public Function UXBeginComment(ByVal reply As Boolean) As Boolean
 UXBeginComment = UXMessageBegin(reply)
End Function
Public Function UXSendComment() As Boolean
 UXSendComment = UXMessagePost()
End Function
Public Function UXRefreshComments() As Boolean
 UXRefreshComments = UXMessageRefresh()
End Function
Public Function CommentsToEditor() As Boolean
 If CurrentProject.AllForms("frmEditor").IsLoaded Then Forms!frmEditor.SetFocus: CommentsToEditor = True
End Function
Public Function CommentsToMyWork() As Boolean
 CommentsToMyWork = UXWorkBack()
End Function
Public Function UXCommentState() As Boolean
 UXCommentState = UXMessageRefresh()
End Function
