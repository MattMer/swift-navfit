Attribute VB_Name = "modReviewComments"
Option Compare Database
Option Explicit
Private Function CommentProfile() As Long
    CommentProfile = Nz(DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1"), 0)
    If CommentProfile = 0 Then Err.Raise vbObjectError + 1700, , "Configure a contributor first."
End Function
Private Sub CommentTarget(ByVal reportID As Long, ByVal revisionID As Long)
    If DCount("*", "ReportRevision", "ReportID=" & reportID & " AND RevisionID=" & revisionID) = 0 Then Err.Raise vbObjectError + 1701, , "Choose a saved revision belonging to this report."
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
        If DCount("*", "ReviewComment", "CommentID=" & threadID & " AND ThreadID Is Null AND ReportID=" & reportID & " AND RevisionID=" & revisionID) = 0 Then Err.Raise vbObjectError + 1704, , "Reply must belong to the selected report and revision thread."
        If ThreadResolved(threadID) Then Err.Raise vbObjectError + 1705, , "Reopen this thread before replying."
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
    If DCount("*", "Report", "ReportID=" & reportID) = 0 Then Err.Raise vbObjectError + 1707, , "Save or select a report first."
    DoCmd.OpenForm "frmReviewComments"
    Forms!frmReviewComments!txtReportID = reportID
    Forms!frmReviewComments!txtCompose = Null
    Call LoadReviewThreads()
    OpenReviewComments = True
End Function
Public Function CommentsFromEditor() As Boolean
    On Error GoTo Failed
    CommentsFromEditor = OpenReviewComments(Nz(Forms!frmEditor!txtReportID, 0)): Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Review comments"
End Function
Public Function LoadReviewThreads() As Boolean
    Dim id As Long
    id = Nz(Forms!frmReviewComments!txtReportID, 0)
    Forms!frmReviewComments!lstThreads.RowSource = "SELECT c.CommentID,r.RevisionNumber AS [Saved revision],p.DisplayName AS Contributor,c.CommentAt AS [Posted at] FROM (ReviewComment AS c INNER JOIN ReportRevision AS r ON c.RevisionID=r.RevisionID) INNER JOIN UserProfile AS p ON c.AuthorProfileID=p.UserProfileID WHERE c.ReportID=" & id & " AND c.ThreadID Is Null ORDER BY c.CommentID DESC"
    Forms!frmReviewComments!lstThreads.Requery
    Forms!frmReviewComments!txtHistory = Null
    Forms!frmReviewComments!lblContext.Caption = UXWorkReport(id) & " | Latest saved revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & CurrentRevision(id)), "?") & ". New threads use saved values; pending edits remain separate."
    LoadReviewThreads = True
End Function
Public Function ShowReviewThread() As Boolean
    Dim r As DAO.Recordset, id As Long, text As String
    id = Nz(Forms!frmReviewComments!lstThreads, 0)
    Set r = CurrentDb.OpenRecordset("SELECT c.*,p.DisplayName FROM ReviewComment AS c INNER JOIN UserProfile AS p ON c.AuthorProfileID=p.UserProfileID WHERE c.ReportID=" & Nz(Forms!frmReviewComments!txtReportID, 0) & " AND (c.CommentID=" & id & " OR c.ThreadID=" & id & ") ORDER BY c.CommentID", dbOpenSnapshot)
    If r.EOF Then r.Close: Forms!frmReviewComments!txtHistory = Null: Exit Function
    If ThreadResolved(id) Then text = "Resolved" Else text = "Open"
    Do While Not r.EOF
        text = text & vbCrLf & vbCrLf & r!DisplayName & " / " & r!CommentAt & " / saved revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & r!RevisionID), "?") & vbCrLf & r!CommentBody
        r.MoveNext
    Loop
    r.Close
    Set r = CurrentDb.OpenRecordset("SELECT e.*,p.DisplayName FROM ReviewCommentEvent AS e INNER JOIN UserProfile AS p ON e.AuthorProfileID=p.UserProfileID WHERE e.ThreadID=" & id & " ORDER BY e.CommentEventID", dbOpenSnapshot)
    Do While Not r.EOF
        text = text & vbCrLf & vbCrLf & r!EventKind & " / " & r!DisplayName & " / " & r!EventAt
        r.MoveNext
    Loop
    r.Close: Forms!frmReviewComments!txtHistory = text: ShowReviewThread = True
End Function
Public Function PostReviewUI(Optional ByVal reply As Boolean = False) As Boolean
    Dim reportID As Long, revisionID As Long, threadID As Long, added As Long
    On Error GoTo Failed
    reportID = Nz(Forms!frmReviewComments!txtReportID, 0)
    If reply Then
        threadID = Nz(Forms!frmReviewComments!lstThreads, 0)
        If threadID = 0 Then Err.Raise vbObjectError + 1706, , "Select a thread before replying."
        revisionID = Nz(DLookup("RevisionID", "ReviewComment", "CommentID=" & threadID), 0)
    Else
        revisionID = Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & reportID), 0)
    End If
    added = AddReviewComment(reportID, revisionID, Nz(Forms!frmReviewComments!txtCompose, ""), threadID)
    Forms!frmReviewComments!txtCompose = Null
    Call LoadReviewThreads()
    If Not reply Then threadID = added
    Forms!frmReviewComments!lstThreads = threadID: Call ShowReviewThread()
    PostReviewUI = True: Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Review comments"
End Function
Public Function ReviewStateUI(ByVal resolved As Boolean) As Boolean
    On Error GoTo Failed
    ReviewStateUI = SetReviewThreadState(Nz(Forms!frmReviewComments!lstThreads, 0), resolved)
    Call ShowReviewThread(): Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Review comments"
End Function
Public Function ReturnFromReviewComments() As Boolean
    Dim id As Long
    id = Nz(Forms!frmReviewComments!txtReportID, 0)
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Nz(Forms!frmEditor!txtReportID, 0) = id Then Forms!frmEditor.SetFocus: ReturnFromReviewComments = True: Exit Function
        If Not UXLeaveEditor() Then Exit Function
    End If
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    ReturnFromReviewComments = True
End Function
