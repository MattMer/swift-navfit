Attribute VB_Name = "modReportMessages"
Option Compare Binary
Option Explicit
'One report-level conversation system. Legacy records remain unchanged and mapped.
Private Function CopyLegacyMessage(ByVal rows As DAO.Recordset, ByVal note As Boolean, ByVal parent As Long) As Long
 Dim target As DAO.Recordset, profile As DAO.Recordset, author As Variant, caption As String
 Set target = CurrentDb.OpenRecordset("ReportMessage", dbOpenDynaset)
 target.AddNew
 If note Then
  target!MessageGuid = rows!NoteGuid: target!LegacyNoteID = rows!NoteID
  author = rows!AuthorProfileID: caption = Nz(rows!ContributorName, "Legacy contributor")
  target!CreatedAt = rows!CreatedAt: target!MessageBody = rows!NoteText
 Else
  target!MessageGuid = rows!CommentGuid: target!LegacyCommentID = rows!CommentID
  target!LegacyRevisionID = rows!RevisionID: author = rows!AuthorProfileID
  Set profile = CurrentDb.OpenRecordset("SELECT DisplayName FROM UserProfile WHERE UserProfileID=" & Nz(author, 0), dbOpenSnapshot)
  caption = "Legacy contributor": If Not profile.EOF Then caption = Nz(profile!DisplayName, caption)
  profile.Close
  target!CreatedAt = rows!CommentAt: target!MessageBody = rows!CommentBody
 End If
 target!ReportID = rows!ReportID: target!ContributorName = caption
 If Not IsNull(author) Then target!AuthorProfileID = author
 If parent > 0 Then target!ParentMessageID = parent
 target.Update: target.Bookmark = target.LastModified: CopyLegacyMessage = target!MessageID: target.Close
End Function
Public Function MigrateReportMessages(Optional ByVal report As Long = 0) As Long
 Dim rows As DAO.Recordset, parents As DAO.Recordset, filter As String, parent As Long, value As Long, w As DAO.Workspace, trans As Boolean, failure As Long, message As String
 On Error GoTo Failed
 Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
 If report > 0 Then filter = " AND ReportID=" & report
 Set rows = CurrentDb.OpenRecordset("SELECT * FROM ReportNote WHERE NoteID NOT IN (SELECT LegacyNoteID FROM ReportMessage WHERE LegacyNoteID Is Not Null)" & filter, dbOpenSnapshot)
 Do While Not rows.EOF
  value = CopyLegacyMessage(rows, True, 0): MigrateReportMessages = MigrateReportMessages + 1: rows.MoveNext
 Loop
 rows.Close
 Set rows = CurrentDb.OpenRecordset("SELECT * FROM ReviewComment WHERE ThreadID Is Null AND CommentID NOT IN (SELECT LegacyCommentID FROM ReportMessage WHERE LegacyCommentID Is Not Null)" & filter, dbOpenSnapshot)
 Do While Not rows.EOF
  value = CopyLegacyMessage(rows, False, 0): MigrateReportMessages = MigrateReportMessages + 1: rows.MoveNext
 Loop
 rows.Close
 Set rows = CurrentDb.OpenRecordset("SELECT * FROM ReviewComment WHERE ThreadID Is Not Null AND CommentID NOT IN (SELECT LegacyCommentID FROM ReportMessage WHERE LegacyCommentID Is Not Null)" & filter, dbOpenSnapshot)
 Do While Not rows.EOF
  Set parents = CurrentDb.OpenRecordset("SELECT MessageID FROM ReportMessage WHERE LegacyCommentID=" & rows!ThreadID & " AND ReportID=" & rows!ReportID, dbOpenSnapshot)
  parent = 0: If Not parents.EOF Then parent = parents!MessageID
  parents.Close
  If parent = 0 Then Err.Raise vbObjectError + 8710, , "Legacy comment parent is missing; conversation conversion stopped."
  value = CopyLegacyMessage(rows, False, parent): MigrateReportMessages = MigrateReportMessages + 1: rows.MoveNext
 Loop
 rows.Close: w.CommitTrans: trans = False: Exit Function
Failed:
 failure = Err.Number: message = Err.Description
 Resume Cleanup
Cleanup:
 On Error Resume Next
 If Not rows Is Nothing Then rows.Close
 If trans Then w.Rollback
 On Error GoTo 0: Err.Raise failure, "MigrateReportMessages", message
End Function
Public Function AddReportMessage(ByVal report As Long, ByVal body As String, Optional ByVal parent As Long = 0) As Long
 Dim rows As DAO.Recordset, author As Long
 author = MajorAuthor()
 If DCount("*", "Report", "ReportID=" & report) <> 1 Then Err.Raise 5, , "Save the report before posting a note."
 If Nz(DLookup("IsArchived", "Report", "ReportID=" & report), True) Then Err.Raise 5, , "Recover the archived report before posting notes."
 If Len(Trim$(Replace(Replace(Replace(body, vbCr, ""), vbLf, ""), vbTab, ""))) = 0 Or Len(body) > 10000 Then Err.Raise 5, , "Enter a note of 1 to 10,000 characters."
 If parent > 0 Then
  If DCount("*", "ReportMessage", "MessageID=" & parent & " AND ReportID=" & report & " AND ParentMessageID Is Null") <> 1 Then Err.Raise 5, , "Choose a conversation belonging to this report."
 End If
 Set rows = CurrentDb.OpenRecordset("ReportMessage", dbOpenDynaset)
 rows.AddNew: rows!MessageGuid = NewIdentity(): rows!ReportID = report
 rows!AuthorProfileID = author: rows!ContributorName = DLookup("DisplayName", "UserProfile", "UserProfileID=" & author)
 rows!CreatedAt = Now(): rows!MessageBody = body
 If parent > 0 Then rows!ParentMessageID = parent
 rows.Update: rows.Bookmark = rows.LastModified: AddReportMessage = rows!MessageID: rows.Close
End Function
Public Function SaveReportMessageDraft(ByVal identity As String, ByVal report As Long, ByVal expected As Long, ByVal body As String, ByVal parent As Long) As Long
 Dim w As DAO.Workspace, failure As Long, message As String
 Set w = DBEngine.Workspaces(0): w.BeginTrans
 On Error GoTo Failed
 If parent > 0 Then
  If DCount("*", "ReportMessage", "MessageID=" & parent & " AND ReportID=" & report & " AND ParentMessageID Is Null") <> 1 Then Err.Raise 5, , "The reply conversation belongs to another report."
 End If
 SaveReportMessageDraft = SaveReportNoteDraft(identity, report, expected, body)
 If parent > 0 Then
  CurrentDb.Execute "UPDATE ReportNoteDraft SET ParentMessageID=" & parent & " WHERE DraftGuid=" & QuoteSQL(identity), dbFailOnError
 Else
  CurrentDb.Execute "UPDATE ReportNoteDraft SET ParentMessageID=Null WHERE DraftGuid=" & QuoteSQL(identity), dbFailOnError
 End If
 w.CommitTrans: Exit Function
Failed:
 failure = Err.Number: message = Err.Description: Resume Cleanup
Cleanup:
 On Error Resume Next: w.Rollback
 On Error GoTo 0: Err.Raise failure, "SaveReportMessageDraft", message
End Function
Public Function PostReportMessageDraft(ByVal identity As String, ByVal expected As Long) As Long
 Dim rows As DAO.Recordset, w As DAO.Workspace, failure As Long, message As String
 Set w = DBEngine.Workspaces(0): w.BeginTrans
 On Error GoTo Failed
 Set rows = CurrentDb.OpenRecordset("SELECT * FROM ReportNoteDraft WHERE DraftGuid=" & QuoteSQL(identity), dbOpenSnapshot)
 If rows.EOF Then Err.Raise 5, , "Draft no longer exists. Refresh your notes."
 If rows!AuthorProfileID <> MajorAuthor() Or rows!VersionCounter <> expected Then Err.Raise 5, , "The note draft changed. Resume its latest version."
 PostReportMessageDraft = AddReportMessage(rows!ReportID, Nz(rows!NoteText, ""), Nz(rows!ParentMessageID, 0))
 rows.Close
 DiscardReportNoteDraft identity, expected
 w.CommitTrans: Exit Function
Failed:
 failure = Err.Number: message = Err.Description: Resume Cleanup
Cleanup:
 On Error Resume Next
 If Not rows Is Nothing Then rows.Close
 w.Rollback
 On Error GoTo 0: Err.Raise failure, "PostReportMessageDraft", message
End Function
