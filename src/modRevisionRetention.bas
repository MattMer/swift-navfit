Attribute VB_Name = "modRevisionRetention"
Option Compare Binary
Option Explicit

'Only new-policy Working states can be replaced. Missing metadata means legacy:
'never prune imported or pre-policy history, even if it is hidden in the editor.
Public Function RevisionRetentionKind(ByVal revision As Long) As String
    RevisionRetentionKind = Nz(MajorScalar("SELECT RetentionKind FROM RevisionRetention WHERE RevisionID=" & revision), "Legacy")
End Function

Public Function RetainedNarrativeKey(ByVal key As String, ByVal kind As String) As Boolean
    RetainedNarrativeKey = (key = "OfficialComments" Or key = "CommandEmployment" Or key = "DutyDescription")
    If kind = "CHIEFEVAL" And key Like "Trait##Comment#" Then RetainedNarrativeKey = True
End Function

Public Sub SetRevisionRetention(ByVal revision As Long, ByVal kind As String, Optional ByVal reason As String = "")
    If kind <> "Working" And kind <> "Checkpoint" And kind <> "Narrative" And kind <> "Protected" Then Err.Raise 5, , "Unknown revision retention kind."
    If PersistenceCount("ReportRevision", "RevisionID=" & revision) <> 1 Then Err.Raise 5, , "Choose an existing report state."
    CurrentDb.Execute "INSERT INTO RevisionRetention (RevisionID,RetentionKind,CheckpointReason) VALUES (" & revision & "," & QuoteSQL(kind) & "," & QuoteSQL(Left$(reason, 255)) & ")", dbFailOnError
End Sub

'Caller owns the workflow transaction: seal exactly the state being handed off.
Public Sub RetainReportCheckpoint(ByVal reportID As Long, ByVal reason As String)
    Dim revision As Long
    revision = CurrentRevision(reportID)
    If RevisionRetentionKind(revision) = "Legacy" Then
        SetRevisionRetention revision, "Checkpoint", reason
    Else
        CurrentDb.Execute "UPDATE RevisionRetention SET RetentionKind='Checkpoint',CheckpointReason=" & QuoteSQL(Left$(reason, 255)) & " WHERE RevisionID=" & revision, dbFailOnError
    End If
    AuditRecord "Revision", revision
End Sub

'Tables that belong to one revision and are deleted with it.
Public Function RevisionOwnedTable(ByVal tableName As String) As Boolean
    Select Case tableName
        Case "RevisionRetention", "RevisionValue", "RevisionChange", "RevisionDate", "RevisionChoice": RevisionOwnedTable = True
    End Select
End Function
'Whether a column points at a revision, so a state it names must never be deleted.
'Deleting history depends on this rule: a column it misses lets a referenced state go.
'The retention test fails when a relationship to ReportRevision.RevisionID, or a
'whole-number column named like a revision, is not recognised here.
Public Function RevisionReferenceField(ByVal tableName As String, ByVal fieldName As String) As Boolean
    If RevisionOwnedTable(tableName) Then Exit Function
    If tableName = "ReportRevision" And fieldName = "RevisionID" Then Exit Function
    RevisionReferenceField = (fieldName Like "*RevisionID" Or fieldName = "RevisionSourceID" Or fieldName = "ExpectedRevision" Or fieldName = "AppliedRevision" Or fieldName = "BaseRevision")
End Function
'References include semantic revision pointers without enforced foreign keys.
'Never rewrite an undo, conversation, signature, import preview or review pointer.
Private Function RevisionHasReferences(ByVal revision As Long) As Boolean
    Dim t As DAO.TableDef, f As DAO.Field, db As DAO.Database
    Set db = CurrentDb
    For Each t In db.TableDefs
        If Left$(t.Name, 4) <> "MSys" And Left$(t.Name, 1) <> "~" Then
            For Each f In t.Fields
                If f.Type = dbLong Then
                    If RevisionReferenceField(t.Name, f.Name) Then
                        If PersistenceCount(t.Name, "[" & f.Name & "]=" & revision) > 0 Then RevisionHasReferences = True: Exit Function
                    End If
                End If
            Next
        End If
    Next
End Function

Public Sub RetireWorkingRevision(ByVal revision As Long, Optional ByVal allowUnreferencedProtected As Boolean = False)
    Dim db As DAO.Database, table As Variant
    If revision <= 0 Then Exit Sub
    If RevisionRetentionKind(revision) <> "Working" Then
        If Not allowUnreferencedProtected Or RevisionRetentionKind(revision) <> "Protected" Then Exit Sub
    End If
    Set db = CurrentDb
    If RevisionHasReferences(revision) Then
        db.Execute "UPDATE RevisionRetention SET RetentionKind='Protected' WHERE RevisionID=" & revision, dbFailOnError
        Exit Sub
    End If
    For Each table In Array("RevisionRetention", "RevisionDate", "RevisionChoice", "RevisionChange", "RevisionValue", "ReportRevision")
        db.Execute "DELETE FROM [" & table & "] WHERE RevisionID=" & revision, dbFailOnError
    Next
End Sub

Public Function HistoryRevisionPredicate(ByVal alias As String) As String
    Dim id As String, narrative As String
    id = alias & ".RevisionID"
    narrative = "FieldKey In ('OfficialComments','CommandEmployment','DutyDescription') Or FieldKey Like 'Trait##Comment#'"
    HistoryRevisionPredicate = "(" & id & " In (SELECT RevisionID FROM RevisionRetention WHERE RetentionKind In ('Checkpoint','Narrative')) OR (" & id & " Not In (SELECT RevisionID FROM RevisionRetention) AND (" & alias & ".ParentRevisionID Is Null"
    HistoryRevisionPredicate = HistoryRevisionPredicate & " OR " & id & " In (SELECT RevisionID FROM RevisionChange WHERE " & narrative & ")"
    HistoryRevisionPredicate = HistoryRevisionPredicate & " OR " & id & " In (SELECT RevisionID FROM RouteRunEvent) OR " & id & " In (SELECT RevisionID FROM RouteRun)"
    HistoryRevisionPredicate = HistoryRevisionPredicate & " OR " & id & " In (SELECT ApprovedRevisionID FROM RouteRun) OR " & id & " In (SELECT IncomingRevisionID FROM RouteRunStep) OR " & id & " In (SELECT OutgoingRevisionID FROM RouteRunStep)"
    HistoryRevisionPredicate = HistoryRevisionPredicate & " OR " & id & " In (SELECT SignedRevisionID FROM ReportSignoff) OR " & id & " In (SELECT RevisionID FROM ReviewTask) OR " & id & " In (SELECT FinishedRevisionID FROM ReviewTask))))"
End Function

Public Function RevisionHistoryLabel(ByVal revision As Long, ByVal reason As Variant) As String
    Dim kind As String, label As String
    kind = RevisionRetentionKind(revision)
    If kind = "Checkpoint" Then
        label = Nz(MajorScalar("SELECT CheckpointReason FROM RevisionRetention WHERE RevisionID=" & revision), "Checkpoint")
        RevisionHistoryLabel = "Checkpoint: " & label
    ElseIf kind = "Narrative" Then
        RevisionHistoryLabel = "Text edit: " & Nz(reason, "")
    Else
        RevisionHistoryLabel = Nz(reason, "")
    End If
End Function

Public Sub AssertRevisionRetention()
    If Nz(MajorScalar("SELECT Count(*) FROM RevisionRetention AS h LEFT JOIN ReportRevision AS r ON h.RevisionID=r.RevisionID WHERE r.RevisionID Is Null"), 0) > 0 Then Err.Raise vbObjectError + 9891, , "Revision retention has an orphaned snapshot pointer."
End Sub
