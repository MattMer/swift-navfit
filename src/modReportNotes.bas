Attribute VB_Name = "modReportNotes"
Option Compare Binary
Option Explicit
Public Function SaveReportNoteDraft(ByVal draftGuid As String, ByVal reportID As Long, ByVal expectedVersion As Long, ByVal body As String) As Long
    Dim r As DAO.Recordset, w As DAO.Workspace, author As Long, failure As Long, message As String
    author = MajorAuthor()
    If Len(body) > 10000 Then Err.Raise vbObjectError + 6950, , "Keep each local note within 10,000 characters."
    If Len(draftGuid) <> 38 Then Err.Raise vbObjectError + 6950, , "Invalid note draft identity."
    If Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & reportID), True) Then Err.Raise vbObjectError + 6950, , "Save or recover the report before adding local notes."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ReportNoteDraft WHERE DraftGuid=" & QuoteSQL(draftGuid), dbOpenDynaset)
    If r.EOF Then
        If expectedVersion <> 0 Then Err.Raise vbObjectError + 6950, , "Note checkpoint was removed. Recover or start another note."
        r.AddNew: r!DraftGuid = draftGuid: r!ReportID = reportID: r!AuthorProfileID = author: r!VersionCounter = 1
    Else
        If r!AuthorProfileID <> author Or r!ReportID <> reportID Then Err.Raise vbObjectError + 6950, , "This note checkpoint belongs to another contributor or report."
        If r!VersionCounter <> expectedVersion Then Err.Raise vbObjectError + 6950, , "Note checkpoint changed in another window. Recover its latest version."
        r.Edit: r!VersionCounter = r!VersionCounter + 1
    End If
    r!NoteText = body: r!CheckpointAt = Now(): r.Update: r.Bookmark = r.LastModified: SaveReportNoteDraft = r!VersionCounter: r.Close
    w.CommitTrans: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next: w.Rollback: If Not r Is Nothing Then r.Close
    On Error GoTo 0: Err.Raise failure, "SaveReportNoteDraft", message
End Function
Public Function PostReportNote(ByVal draftGuid As String, ByVal expectedVersion As Long) As Long
    Dim r As DAO.Recordset, note As DAO.Recordset, w As DAO.Workspace, author As Long, op As Long, failure As Long, message As String, id As Long
    author = MajorAuthor(): Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ReportNoteDraft WHERE DraftGuid=" & QuoteSQL(draftGuid), dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 6951, , "Checkpoint the note before posting."
    If r!AuthorProfileID <> author Or r!VersionCounter <> expectedVersion Then Err.Raise vbObjectError + 6951, , "Contributor or note checkpoint changed. Recover before posting."
    id = r!ReportID
    If Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & id), True) Then Err.Raise vbObjectError + 6951, , "Recover the archived report before posting a note."
    If Len(Trim$(Nz(r!NoteText, ""))) = 0 Then Err.Raise vbObjectError + 6951, , "Enter a note before posting."
    Set note = CurrentDb.OpenRecordset("ReportNote", dbOpenDynaset)
    note.AddNew: note!NoteGuid = NewIdentity(): note!ReportID = id: note!AuthorProfileID = author
    note!ContributorName = MajorScalar("SELECT DisplayName FROM UserProfile WHERE UserProfileID=" & author)
    note!NoteText = r!NoteText: note!CreatedAt = Now(): note!SourceKind = "Contributor"
    note.Update: note.Bookmark = note.LastModified: PostReportNote = note!NoteID: note.Close: r.Close
    CurrentDb.Execute "DELETE FROM ReportNoteDraft WHERE DraftGuid=" & QuoteSQL(draftGuid), dbFailOnError
    op = BeginOperation("Post local report note", "Report " & id & "; note " & PostReportNote)
    FinishOperation op, "Completed", "Contributor attribution recorded. Official content and saved revisions unchanged."
    w.CommitTrans: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next: w.Rollback: If Not r Is Nothing Then r.Close
    If Not note Is Nothing Then note.Close
    On Error GoTo 0: Err.Raise failure, "PostReportNote", message
End Function
Public Sub DiscardReportNoteDraft(ByVal draftGuid As String, ByVal expectedVersion As Long)
    Dim q As DAO.QueryDef
    Set q = CurrentDb.CreateQueryDef("", "PARAMETERS identity Text(38), author Long, version Long; DELETE FROM ReportNoteDraft WHERE DraftGuid=identity AND AuthorProfileID=author AND VersionCounter=version")
    q!identity = draftGuid: q!author = MajorAuthor(): q!version = expectedVersion: q.Execute dbFailOnError
    If q.RecordsAffected <> 1 Then q.Close: Err.Raise vbObjectError + 6952, , "Note checkpoint changed; discard was not applied."
    q.Close
End Sub
Public Function MigrateLegacyReportNotes() As Long
    Dim r As DAO.Recordset, note As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT ReportID,DraftNotes,ModifiedAt FROM Report WHERE DraftNotes Is Not Null AND DraftNotes<>'' AND ReportID NOT IN (SELECT LegacyReportID FROM ReportNote WHERE LegacyReportID Is Not Null)", dbOpenSnapshot)
    Set note = CurrentDb.OpenRecordset("ReportNote", dbOpenDynaset)
    Do While Not r.EOF
        note.AddNew: note!NoteGuid = NewIdentity(): note!ReportID = r!ReportID: note!LegacyReportID = r!ReportID
        note!ContributorName = "Legacy notes (author not tracked)": note!NoteText = r!DraftNotes: note!CreatedAt = r!ModifiedAt: note!SourceKind = "Legacy"
        note.Update: MigrateLegacyReportNotes = MigrateLegacyReportNotes + 1: r.MoveNext
    Loop
    r.Close: note.Close
End Function
