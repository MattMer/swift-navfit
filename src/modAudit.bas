Attribute VB_Name = "modAudit"
Option Compare Binary
Option Explicit

Private Sub AuditRegisterActor(ByVal id As Long)
    Dim r As DAO.Recordset, p As DAO.Recordset
    If PersistenceCount("AuditActor", "ActorID=" & id) > 0 Then
        If CStr(MajorScalar("SELECT ProfileGuid FROM AuditActor WHERE ActorID=" & id)) <> CStr(MajorScalar("SELECT ProfileGuid FROM UserProfile WHERE UserProfileID=" & id)) Then Err.Raise vbObjectError + 9052, "modAudit.AuditRegisterActor", "Contributor identity differs from retained audit identity. History was preserved."
        Exit Sub
    End If
    Set p = CurrentDb.OpenRecordset("SELECT ProfileGuid,DisplayName FROM UserProfile WHERE UserProfileID=" & id, dbOpenSnapshot)
    If p.EOF Then Err.Raise vbObjectError + 9052, "modAudit.AuditRegisterActor", "Audit actor identity is missing."
    Set r = CurrentDb.OpenRecordset("AuditActor", dbOpenDynaset)
    r.AddNew: r!ActorID = id: r!ProfileGuid = p!ProfileGuid: r!RegisteredName = p!DisplayName: r!RegisteredAt = Now(): r.Update
    r.Close: p.Close
End Sub

'Canonical timeline envelopes link to domain evidence, never replace it.
'No ambient operation global: callers pass correlation or it stays independent.
Public Function AuditRecord(ByVal source As String, ByVal id As Long, Optional ByVal operation As Long = 0, Optional ByVal historical As Boolean = False, Optional ByVal outcome As Boolean = False) As Long
    Dim db As DAO.Database, r As DAO.Recordset, e As DAO.Recordset, sql As String, pointer As String
    Dim report As Variant, revision As Variant, author As Variant, stamp As Variant, label As String, detail As Variant
    Dim guid As String, version As Long, existing As Variant, task As Variant, oldFolder As Variant, newFolder As Variant, caption As Variant
    If source = "Revision" Then
        If RevisionRetentionKind(id) = "Working" Then
            If operation <= 0 Then Exit Function
            CurrentDb.Execute "UPDATE RevisionRetention SET RetentionKind='Protected' WHERE RevisionID=" & id, dbFailOnError
        End If
    End If
    Set db = CurrentDb
    report = Null: revision = Null: author = Null: stamp = Null: task = Null: oldFolder = Null: newFolder = Null: caption = Null: detail = Null
    Select Case source
        Case "Revision": pointer = "RevisionSourceID": sql = "SELECT * FROM ReportRevision WHERE RevisionID=" & id
        Case "Report event": pointer = "ReportEventSourceID": sql = "SELECT * FROM ReportEvent WHERE ReportEventID=" & id
        Case "Folder move": pointer = "FolderMoveSourceID": sql = "SELECT * FROM FolderMove WHERE FolderMoveID=" & id
        Case "Comment event": pointer = "CommentEventSourceID": sql = "SELECT e.*,c.ReportID,c.RevisionID FROM ReviewCommentEvent AS e INNER JOIN ReviewComment AS c ON e.ThreadID=c.CommentID WHERE e.CommentEventID=" & id
        Case "Review event": pointer = "TaskEventSourceID": sql = "SELECT e.*,t.ReportID,t.RevisionID FROM ReviewTaskEvent AS e INNER JOIN ReviewTask AS t ON e.TaskID=t.TaskID WHERE e.TaskEventID=" & id
        Case "Operation": pointer = "JournalSourceID": sql = "SELECT * FROM OperationJournal WHERE OperationID=" & id: operation = id
        Case "Message": pointer = "MessageSourceID": sql = "SELECT * FROM ReportMessage WHERE MessageID=" & id
        Case Else: Err.Raise vbObjectError + 9050, "modAudit.AuditRecord", "Unknown audit evidence type."
    End Select
    If outcome Then version = Nz(MajorScalar("SELECT Max(SourceVersion) FROM AuditEvent WHERE " & pointer & "=" & id), 0) + 1
    existing = MajorScalar("SELECT AuditID FROM AuditEvent WHERE " & pointer & "=" & id & " AND SourceVersion=" & version)
    If Not IsNull(existing) Then AuditRecord = CLng(existing): Exit Function
    Set r = db.OpenRecordset(sql, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 9051, "modAudit.AuditRecord", "Audit evidence is missing."
    Select Case source
        Case "Revision"
            report = r!ReportID: revision = id: author = r!AuthorProfileID: stamp = r!CommittedAt: label = "Revision saved": detail = r!Reason
            If operation = 0 Then operation = Nz(MajorScalar("SELECT Min(OperationID) FROM ImportReceipt WHERE AppliedRevisionID=" & id & " AND (BeforeRevisionID Is Null OR BeforeRevisionID<>AppliedRevisionID) HAVING Min(OperationID)=Max(OperationID)"), 0)
        Case "Report event"
            report = r!ReportID: author = r!AuthorProfileID: stamp = r!EventAt: label = Nz(r!EventKind, "Report event")
        Case "Folder move"
            report = r!ReportID: author = r!AuthorProfileID: stamp = r!MovedAt: label = "Folder moved": oldFolder = r!OldFolderID: newFolder = r!NewFolderID
            If operation = 0 Then operation = Nz(MajorScalar("SELECT Min(OperationID) FROM FolderBatchChange WHERE AppliedMoveID=" & id & " HAVING Min(OperationID)=Max(OperationID)"), 0)
        Case "Comment event", "Review event"
            report = r!ReportID: revision = r!RevisionID: author = r!AuthorProfileID: stamp = r!EventAt: label = Nz(r!EventKind, source)
            If source = "Review event" Then task = r!TaskID: detail = r!Notes
        Case "Operation"
            author = r!AuthorProfileID: stamp = r!StartedAt: label = "Started: " & Nz(r!OperationKind, "Operation"): detail = r!SourceDescription
            If outcome Then author = MajorAuthor(): stamp = r!FinishedAt: label = r!OperationState: detail = r!Outcome
        Case "Message"
            report = r!ReportID: author = r!AuthorProfileID: stamp = r!CreatedAt: caption = r!ContributorName
            revision = r!ContextRevisionID: If IsNull(revision) Then revision = r!LegacyRevisionID
            label = IIf(IsNull(r!ParentMessageID), "Conversation posted", "Reply posted")
    End Select
    r.Close
    If Not IsNull(author) Then AuditRegisterActor CLng(author)
    If Not historical And IsNull(caption) And Not IsNull(author) Then caption = MajorScalar("SELECT DisplayName FROM UserProfile WHERE UserProfileID=" & CLng(author))
    guid = NewIdentity()
    If operation > 0 Then guid = CStr(MajorScalar("SELECT OperationGuid FROM OperationJournal WHERE OperationID=" & operation))
    Set e = db.OpenRecordset("AuditEvent", dbOpenDynaset)
    e.AddNew: e!AuditGuid = NewIdentity(): e!SourceKind = source: e.Fields(pointer) = id: e!SourceVersion = version
    e!ActorID = author: e!ActorCaption = caption: e!EventAt = stamp: e!EventKind = Left$(label, 80): e!EventDetail = detail
    e!ReportID = report: e!RevisionID = revision: e!TaskID = task: e!OldFolderID = oldFolder: e!NewFolderID = newFolder
    e!CorrelationGuid = guid: e!IsHistorical = historical
    If operation > 0 Then e!OperationID = operation
    e.Update: e.Bookmark = e.LastModified: AuditRecord = e!AuditID: e.Close
End Function

Public Sub AuditCorrelateRevision(ByVal revision As Long, ByVal operation As Long)
    Dim db As DAO.Database
    Set db = CurrentDb
    If operation <= 0 Then Exit Sub
    If RevisionRetentionKind(revision) = "Working" Then
        db.Execute "UPDATE RevisionRetention SET RetentionKind='Protected' WHERE RevisionID=" & revision, dbFailOnError
        AuditRecord "Revision", revision, operation
    End If
    db.Execute "UPDATE AuditEvent SET OperationID=" & operation & ",CorrelationGuid=" & QuoteSQL(CStr(MajorScalar("SELECT OperationGuid FROM OperationJournal WHERE OperationID=" & operation))) & " WHERE RevisionSourceID=" & revision & " AND OperationID Is Null", dbFailOnError
End Sub

Public Sub MigrateAuditHistory()
    Dim kinds As Variant, tables As Variant, ids As Variant, i As Long, r As DAO.Recordset, pointer As String
    kinds = Array("Operation", "Revision", "Report event", "Folder move", "Comment event", "Review event", "Message")
    tables = Array("OperationJournal", "ReportRevision", "ReportEvent", "FolderMove", "ReviewCommentEvent", "ReviewTaskEvent", "ReportMessage")
    ids = Array("OperationID", "RevisionID", "ReportEventID", "FolderMoveID", "CommentEventID", "TaskEventID", "MessageID")
    For i = 0 To UBound(kinds)
        Set r = CurrentDb.OpenRecordset("SELECT [" & ids(i) & "] FROM [" & tables(i) & "] ORDER BY [" & ids(i) & "]", dbOpenSnapshot)
        Do While Not r.EOF: AuditRecord CStr(kinds(i)), CLng(r.Fields(0).Value), 0, True: r.MoveNext: Loop
        r.Close
    Next
End Sub

Public Function AuditDetails() As Boolean
    Dim r As DAO.Recordset, id As Long
    id = Nz(Forms!frmOperations!lstAudit, 0)
    Set r = CurrentDb.OpenRecordset("SELECT * FROM AuditEvent WHERE AuditID=" & id, dbOpenSnapshot)
    If Not r.EOF Then Forms!frmOperations!txtOutcome = r!EventKind & vbCrLf & "Event " & r!AuditGuid & vbCrLf & "Correlation " & r!CorrelationGuid & vbCrLf & Nz(r!EventDetail, "See linked domain history.") & vbCrLf & IIf(r!IsHistorical, "Historical envelope; original evidence retained. Unrecorded context remains unknown.", "Recorded with the domain change.")
    r.Close: AuditDetails = True
End Function
