Attribute VB_Name = "modReportNotes"
Option Compare Binary
Option Explicit
Public Function SaveReportNoteDraft(ByVal draftGuid As String, ByVal reportID As Long, ByVal expectedVersion As Long, ByVal body As String) As Long
    Dim r As DAO.Recordset, w As DAO.Workspace, author As Long, failure As Long, message As String
    author = MajorAuthor()
    If Len(body) > 10000 Then Err.Raise vbObjectError + 6950, "modReportNotes.SaveReportNoteDraft", "Keep each local note within 10,000 characters."
    If Len(draftGuid) <> 38 Then Err.Raise vbObjectError + 6950, "modReportNotes.SaveReportNoteDraft", "Invalid note draft identity."
    If Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & reportID), True) Then Err.Raise vbObjectError + 6950, "modReportNotes.SaveReportNoteDraft", "Save or recover the report before adding local notes."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ReportNoteDraft WHERE DraftGuid=" & QuoteSQL(draftGuid), dbOpenDynaset)
    If r.EOF Then
        If expectedVersion <> 0 Then Err.Raise vbObjectError + 6950, "modReportNotes.SaveReportNoteDraft", "Note checkpoint was removed. Recover or start another note."
        r.AddNew: r!DraftGuid = draftGuid: r!ReportID = reportID: r!AuthorProfileID = author: r!VersionCounter = 1
    Else
        If r!AuthorProfileID <> author Or r!ReportID <> reportID Then Err.Raise vbObjectError + 6950, "modReportNotes.SaveReportNoteDraft", "This note checkpoint belongs to another contributor or report."
        If r!VersionCounter <> expectedVersion Then Err.Raise vbObjectError + 6950, "modReportNotes.SaveReportNoteDraft", "Note checkpoint changed in another window. Recover its latest version."
        r.Edit: r!VersionCounter = r!VersionCounter + 1
    End If
    r!NoteText = body: r!CheckpointAt = Now(): r.Update: r.Bookmark = r.LastModified: SaveReportNoteDraft = r!VersionCounter: r.Close
    w.CommitTrans: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modReportNotes", "SaveReportNoteDraft", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary: If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function
Public Function PostReportNote(ByVal draftGuid As String, ByVal expectedVersion As Long) As Long
    'Compatibility result is a negative canonical message token, never a NoteID.
    PostReportNote = -PostReportMessageDraft(draftGuid, expectedVersion)
End Function
Public Sub DiscardReportNoteDraft(ByVal draftGuid As String, ByVal expectedVersion As Long)
    Dim q As DAO.QueryDef
    Set q = CurrentDb.CreateQueryDef("", "PARAMETERS identity Text(38), author Long, version Long; DELETE FROM ReportNoteDraft WHERE DraftGuid=identity AND AuthorProfileID=author AND VersionCounter=version")
    q!identity = draftGuid: q!author = MajorAuthor(): q!version = expectedVersion: q.Execute dbFailOnError
    If q.RecordsAffected <> 1 Then q.Close: Err.Raise vbObjectError + 6952, "modReportNotes.DiscardReportNoteDraft", "Note checkpoint changed; discard was not applied."
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
