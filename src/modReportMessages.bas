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
 Dim rows As DAO.Recordset, parents As DAO.Recordset, filter As String, parent As Long, value As Long, w As DAO.Workspace, trans As Boolean, failure As Long, message As String, created As Object, mapped As Long
 On Error GoTo Failed
 Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
 Set created = CreateObject("Scripting.Dictionary")
 If report > 0 Then filter = " AND ReportID=" & report
 Set rows = CurrentDb.OpenRecordset("SELECT * FROM ReportNote WHERE NoteID NOT IN (SELECT LegacyNoteID FROM ReportMessage WHERE LegacyNoteID Is Not Null)" & filter & " ORDER BY NoteID", dbOpenSnapshot)
 Do While Not rows.EOF
  value = CopyLegacyMessage(rows, True, 0): MigrateReportMessages = MigrateReportMessages + 1: rows.MoveNext
 Loop
 rows.Close
 'Allocate every identity before linking any parents, including nested/later IDs.
 Set rows = CurrentDb.OpenRecordset("SELECT * FROM ReviewComment WHERE CommentID NOT IN (SELECT LegacyCommentID FROM ReportMessage WHERE LegacyCommentID Is Not Null)" & filter & " ORDER BY CommentID", dbOpenSnapshot)
 Do While Not rows.EOF
  value = CopyLegacyMessage(rows, False, 0): created(CStr(value)) = True: MigrateReportMessages = MigrateReportMessages + 1: rows.MoveNext
 Loop
 rows.Close
 Set rows = CurrentDb.OpenRecordset("SELECT * FROM ReviewComment WHERE ThreadID Is Not Null" & filter, dbOpenSnapshot)
 Do While Not rows.EOF
  parent = Nz(MajorScalar("SELECT MessageID FROM ReportMessage WHERE LegacyCommentID=" & rows!ThreadID & " AND ReportID=" & rows!ReportID), 0)
  mapped = Nz(MajorScalar("SELECT MessageID FROM ReportMessage WHERE LegacyCommentID=" & rows!CommentID & " AND ReportID=" & rows!ReportID), 0)
  If parent = 0 Or mapped = 0 Then Err.Raise vbObjectError + 8710, , "Legacy parent is missing or belongs to another report. Conversion rolled back."
  If created.Exists(CStr(mapped)) Then
   CurrentDb.Execute "UPDATE ReportMessage SET ParentMessageID=" & parent & " WHERE MessageID=" & mapped, dbFailOnError
  ElseIf Nz(MajorScalar("SELECT ParentMessageID FROM ReportMessage WHERE MessageID=" & mapped), 0) <> parent Then
   Err.Raise vbObjectError + 8710, , "Existing legacy conversation mapping differs. Original history was preserved."
  End If
  rows.MoveNext
 Loop
 rows.Close
 Set rows = CurrentDb.OpenRecordset("SELECT MessageID,ReportID FROM ReportMessage WHERE (LegacyNoteID Is Not Null OR LegacyCommentID Is Not Null)" & filter, dbOpenSnapshot)
 Do While Not rows.EOF
  value = ReportMessageRoot(CLng(rows!MessageID), CLng(rows!ReportID))
  If SCHEMA_VERSION >= 29 Then AuditRecord "Message", CLng(rows!MessageID), 0, True
  rows.MoveNext
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

Public Function ReportMessageRoot(ByVal id As Long, ByVal report As Long) As Long
 Dim seen As Object, r As DAO.Recordset, parent As Long
 Set seen = CreateObject("Scripting.Dictionary")
 Do While id > 0
  If seen.Exists(CStr(id)) Then Err.Raise vbObjectError + 8711, , "Conversation has a parent cycle. Original history was preserved."
  seen(CStr(id)) = True
  Set r = CurrentDb.OpenRecordset("SELECT ParentMessageID FROM ReportMessage WHERE MessageID=" & id & " AND ReportID=" & report, dbOpenSnapshot)
  If r.EOF Then Err.Raise vbObjectError + 8711, , "Conversation parent is missing or belongs to another report."
  parent = Nz(r!ParentMessageID, 0): r.Close
  If parent = 0 Then ReportMessageRoot = id: Exit Function
  id = parent
 Loop
End Function

Public Function ReportMessageParents(ByVal report As Long) As Object
 Dim map As Object, rows As DAO.Recordset
 Set map = CreateObject("Scripting.Dictionary")
 Set rows = CurrentDb.OpenRecordset("SELECT MessageID,ParentMessageID FROM ReportMessage WHERE ReportID=" & report, dbOpenSnapshot)
 Do While Not rows.EOF: map(CStr(rows!MessageID)) = Nz(rows!ParentMessageID, 0): rows.MoveNext: Loop
 rows.Close: Set ReportMessageParents = map
End Function
Public Function ReportMessageMapRoot(ByVal id As Long, ByVal parents As Object) As Long
 Dim seen As Object, parent As Long
 Set seen = CreateObject("Scripting.Dictionary")
 Do While id > 0
  If seen.Exists(CStr(id)) Or Not parents.Exists(CStr(id)) Then Err.Raise vbObjectError + 8711, , "Conversation parent is cyclic, missing or belongs to another report."
  seen(CStr(id)) = True: parent = CLng(parents(CStr(id)))
  If parent = 0 Then ReportMessageMapRoot = id: Exit Function
  id = parent
 Loop
End Function
Public Function AddReportMessage(ByVal report As Long, ByVal body As String, Optional ByVal parent As Long = 0, Optional ByVal contextRevision As Long = 0) As Long
 Dim rows As DAO.Recordset, author As Long, w As DAO.Workspace, failure As Long, explanation As String
 author = MajorAuthor()
 If PersistenceCount("Report", "ReportID=" & report) <> 1 Then Err.Raise 5, , "Save the report before posting a note."
 If Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & report), True) Then Err.Raise 5, , "Recover the archived report before posting notes."
 If Len(Trim$(Replace(Replace(Replace(body, vbCr, ""), vbLf, ""), vbTab, ""))) = 0 Or Len(body) > 10000 Then Err.Raise 5, , "Enter a note of 1 to 10,000 characters."
 If parent > 0 Then
  If PersistenceCount("ReportMessage", "MessageID=" & parent & " AND ReportID=" & report & " AND ParentMessageID Is Null") <> 1 Then Err.Raise 5, , "Choose a conversation belonging to this report."
 End If
 If contextRevision > 0 Then
  If PersistenceCount("ReportRevision", "RevisionID=" & contextRevision & " AND ReportID=" & report) <> 1 Then Err.Raise 5, , "Revision context belongs to another report."
 End If
 Set w = DBEngine.Workspaces(0): w.BeginTrans
 On Error GoTo Failed
 Set rows = CurrentDb.OpenRecordset("ReportMessage", dbOpenDynaset)
 rows.AddNew: rows!MessageGuid = NewIdentity(): rows!ReportID = report
 rows!AuthorProfileID = author: rows!ContributorName = MajorScalar("SELECT DisplayName FROM UserProfile WHERE UserProfileID=" & author)
 rows!CreatedAt = Now(): rows!MessageBody = body
 If parent > 0 Then rows!ParentMessageID = parent
 If contextRevision > 0 Then rows!ContextRevisionID = contextRevision
 rows.Update: rows.Bookmark = rows.LastModified: AddReportMessage = rows!MessageID: rows.Close
 AuditRecord "Message", AddReportMessage
 w.CommitTrans: Exit Function
Failed:
 failure = Err.Number: explanation = Err.Description
 On Error Resume Next: w.Rollback: On Error GoTo 0
 Err.Raise failure, "AddReportMessage", explanation
End Function
Public Function SaveReportMessageDraft(ByVal identity As String, ByVal report As Long, ByVal expected As Long, ByVal body As String, ByVal parent As Long) As Long
 Dim w As DAO.Workspace, failure As Long, message As String
 Set w = DBEngine.Workspaces(0): w.BeginTrans
 On Error GoTo Failed
 If parent > 0 Then
  If PersistenceCount("ReportMessage", "MessageID=" & parent & " AND ReportID=" & report & " AND ParentMessageID Is Null") <> 1 Then Err.Raise 5, , "The reply conversation belongs to another report."
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
