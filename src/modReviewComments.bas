Attribute VB_Name = "modReviewComments"
Option Compare Database
Option Explicit
'Legacy positive identifiers refer only to retained ReviewComment rows.
'New adapter results are negative canonical message tokens; namespaces never mix.
Public Function AddReviewComment(ByVal reportID As Long, ByVal revisionID As Long, ByVal body As String, Optional ByVal threadID As Long = 0) As Long
    Dim parent As Long
    If PersistenceCount("ReportRevision", "RevisionID=" & revisionID & " AND ReportID=" & reportID) <> 1 Then Err.Raise vbObjectError + 1701, , "Choose a revision belonging to this report."
    If threadID < 0 Then
        parent = -threadID
    ElseIf threadID > 0 Then
        MigrateReportMessages reportID
        parent = Nz(MajorScalar("SELECT MessageID FROM ReportMessage WHERE LegacyCommentID=" & threadID & " AND ReportID=" & reportID), 0)
        If parent = 0 Then Err.Raise vbObjectError + 1704, , "Legacy conversation does not belong to this report."
    End If
    AddReviewComment = -AddReportMessage(reportID, body, parent, revisionID)
End Function
Public Function SetReviewThreadState(ByVal threadID As Long, ByVal resolved As Boolean) As Boolean
    Err.Raise vbObjectError + 1706, , "Conversation resolve/reopen is retired. Notes and replies are independent of approval."
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
