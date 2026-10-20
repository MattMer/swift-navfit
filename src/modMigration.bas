Attribute VB_Name = "modMigration"
Option Compare Database
Option Explicit

Public Function MigrateWorkspace(ByVal sourcePath As String, Optional ByVal showProgress As Boolean = False, Optional ByVal replaceSetup As Boolean = False, Optional ByVal expectedHash As String = "") As String
    Dim s As DAO.Database, t As DAO.Database, w As DAO.Workspace, reader As DAO.Workspace
    Dim r As DAO.Recordset, a As DAO.Recordset, b As DAO.Recordset, q As DAO.QueryDef
    Dim tables As Variant, n As Variant, i As Long, ver As Long, sql As String
    Dim names As String, vals As String, trans As Boolean, message As String, stage As String, sourceSQL As String, temporaryQuery As String
    Dim av As Variant, bv As Variant, fieldName As String, deferred As Variant, relationPair As Variant, copyStarted As Single, verifyStarted As Single
    Dim table As DAO.TableDef, known As Boolean, targetVersion As Long, copiedRows As Long, checkedRows As Long
    On Error GoTo Failed
    Set reader = DBEngine.CreateWorkspace("SwiftRead", "admin", "", dbUseJet)
    Set s = reader.OpenDatabase(sourcePath, False, True)
    Set t = CurrentDb: Set w = DBEngine.Workspaces(0)
    Set r = s.OpenRecordset("AppMetadata", dbOpenSnapshot): ver = r!SchemaVersion: r.Close
    If ver < 1 Or ver > 26 Then Err.Raise vbObjectError + 301, , "Unsupported source schema."
    targetVersion = DLookup("SchemaVersion", "AppMetadata")
    If targetVersion <> 21 And targetVersion <> 22 And targetVersion <> 23 And targetVersion <> 24 And targetVersion <> 25 And targetVersion <> 26 Then Err.Raise vbObjectError + 302, , "Target schema must be 21, 22, 23, 24, 25 or 26."
    If ver > targetVersion Then Err.Raise vbObjectError + 302, , "Migration cannot downgrade the source schema."
    tables = Array("UserProfile", "Workspace", "Folder", "Report", "ReportRevision", "RevisionValue", "RevisionChange", "ReportEvent", "FolderMove", "OperationJournal", "ImportReceipt", "ImportRawValue", "RecoveryDraft", "NarrativeSnippet", "LibraryPreference", "FolderBatchChange", "ReviewComment", "ReviewCommentEvent", "ReviewTask", "ReviewTaskEvent", _
        "SummaryGroup", "GroupMember", "GroupSnapshot", "GroupSnapshotMember", "FieldDefault", "ChangeBatch", "ChangeItem", "RouteTemplate", "RouteTemplateStep", "RouteRun", "RouteRunStep", "ComparisonSet", "ComparisonMember", "MeetingProposal", "RulePack", "RuleDefinition", "ValidationRun", "ValidationFinding", "ReportingSeniorRecord", "ReportingSeniorValue", "RSCABaseline", "GroupRSCABaseline")
    If targetVersion >= 22 Then tables = Split(Join(tables, "|") & "|ReportNote|ReportNoteDraft", "|")
    If targetVersion >= 23 Then tables = Split(Join(tables, "|") & "|CommandMember|MemberReport|PriorEvaluation|PriorEvaluationTrait", "|")
    If targetVersion >= 26 Then tables = Split(Join(tables, "|") & "|ReportMessage", "|")
    For Each table In s.TableDefs
        If Left$(table.Name, 4) <> "MSys" Then
            known = (table.Name = "AppMetadata")
            For Each n In tables
                If table.Name = n Then known = True
            Next
            If Not known Then Err.Raise vbObjectError + 309, , "Unrecognized source table: " & table.Name & ". Migration stopped to prevent data loss."
        End If
    Next
    'One inexpensive guard; TOP 1 never counts all records in a populated table.
    For Each table In t.TableDefs
        If Left$(table.Name, 4) <> "MSys" And table.Name <> "AppMetadata" Then
            If Not (replaceSetup And (table.Name = "Workspace" Or table.Name = "UserProfile")) Then
                Set r = t.OpenRecordset("SELECT TOP 1 * FROM [" & table.Name & "]", dbOpenSnapshot)
                If Not r.EOF Then r.Close: Err.Raise vbObjectError + 303, , "Target must be blank."
                r.Close
            End If
        End If
    Next
    'Saved query metadata must be created outside the data transaction. Otherwise
    'an action-query failure can invalidate DAO handles while rolling back DDL.
    deferred = Array("Report|CurrentRevisionID|ReportID", "ReportRevision|ParentRevisionID|RevisionID", "Folder|ParentFolderID|FolderID", "ReviewComment|ThreadID|CommentID")
    If ver >= 26 Then deferred = Split(Join(deferred, "~") & "~ReportMessage|ParentMessageID|MessageID~ReportNoteDraft|ParentMessageID|DraftGuid", "~")
    temporaryQuery = "zSwiftMigration" & Replace(Replace(Replace(NewIdentity(), "{", ""), "}", ""), "-", "")
    sourceSQL = " IN " & Chr$(34) & sourcePath & Chr$(34)
    For Each relationPair In deferred
        relationPair = Split(CStr(relationPair), "|")
        If Not (ver < 2 And relationPair(0) = "Folder") And Not (ver < 7 And relationPair(0) = "ReviewComment") Then
            Set q = t.CreateQueryDef(temporaryQuery & relationPair(0), "SELECT * FROM [" & relationPair(0) & "]" & sourceSQL)
            q.Close
        End If
    Next
    w.BeginTrans: trans = True
    If replaceSetup Then
        t.Execute "DELETE FROM Workspace", dbFailOnError
        t.Execute "DELETE FROM UserProfile", dbFailOnError
    End If
    copyStarted = Timer
    For Each n In tables
        If ver < 26 And n = "ReportMessage" Then GoTo NextTable
        If ver < 23 And (n = "CommandMember" Or n = "MemberReport" Or n = "PriorEvaluation" Or n = "PriorEvaluationTrait") Then GoTo NextTable
        If ver < 22 And (n = "ReportNote" Or n = "ReportNoteDraft") Then GoTo NextTable
        If ver < 21 And IsRSTable(CStr(n)) Then GoTo NextTable
        If ver < 20 And IsMajorTable(CStr(n)) Then GoTo NextTable
        If ver < 8 And (n = "ReviewTask" Or n = "ReviewTaskEvent") Then GoTo NextTable
        If ver < 7 And (n = "ReviewComment" Or n = "ReviewCommentEvent") Then GoTo NextTable
        If ver < 6 And (n = "OperationJournal" Or n = "NarrativeSnippet" Or n = "LibraryPreference" Or n = "FolderBatchChange") Then GoTo NextTable
        If ver < 5 And n = "RecoveryDraft" Then GoTo NextTable
        If ver < 4 And (n = "ImportReceipt" Or n = "ImportRawValue") Then GoTo NextTable
        If ver = 1 And (n = "Folder" Or n = "FolderMove") Then GoTo NextTable
        copiedRows = 0
        If showProgress Then Application.Run "WorkspaceImportProgress", "Copying " & n & "..."
        stage = "Copy " & n
        names = "": vals = ""
        For i = 0 To s.TableDefs(n).Fields.Count - 1
            fieldName = s.TableDefs(n).Fields(i).Name
            If Len(names) > 0 Then names = names & ",": vals = vals & ","
            names = names & "[" & fieldName & "]"
            If DeferredMigrationField(CStr(n), fieldName) Then
                vals = vals & "Null AS [" & fieldName & "]"
            Else
                vals = vals & "[" & fieldName & "]"
            End If
        Next
        'Explicit source columns preserve IDs and allow newer target-only defaults.
        t.Execute "INSERT INTO [" & n & "] (" & names & ") SELECT " & vals & " FROM [" & n & "]" & sourceSQL, dbFailOnError
        copiedRows = t.RecordsAffected
        If showProgress Then Application.Run "WorkspaceImportProgress", "Copied " & n & ": " & copiedRows & " records copied in bulk"
NextTable:
    Next
    If showProgress Then Application.Run "WorkspaceImportProgress", "Restoring report revision and folder relationships..."
    For Each relationPair In deferred
        relationPair = Split(CStr(relationPair), "|")
        If Not (ver < 2 And relationPair(0) = "Folder") And Not (ver < 7 And relationPair(0) = "ReviewComment") Then
            stage = "Restore " & relationPair(0) & "." & relationPair(1)
            t.Execute "UPDATE [" & relationPair(0) & "] AS dst INNER JOIN [" & temporaryQuery & relationPair(0) & "] AS src ON dst.[" & relationPair(2) & "]=src.[" & relationPair(2) & "] SET dst.[" & relationPair(1) & "]=src.[" & relationPair(1) & "]", dbFailOnError
        End If
    Next
    If showProgress Then Application.Run "WorkspaceImportProgress", "Bulk copying finished in " & Format$(MigrationSeconds(copyStarted), "0.00") & " seconds. Verifying copied values..."
    verifyStarted = Timer
    For Each n In tables
        If ver < 26 And n = "ReportMessage" Then GoTo NextCheck
        If ver < 23 And (n = "CommandMember" Or n = "MemberReport" Or n = "PriorEvaluation" Or n = "PriorEvaluationTrait") Then GoTo NextCheck
        If ver < 22 And (n = "ReportNote" Or n = "ReportNoteDraft") Then GoTo NextCheck
        If ver < 21 And IsRSTable(CStr(n)) Then GoTo NextCheck
        If ver < 20 And IsMajorTable(CStr(n)) Then GoTo NextCheck
        If ver < 8 And (n = "ReviewTask" Or n = "ReviewTaskEvent") Then GoTo NextCheck
        If ver < 7 And (n = "ReviewComment" Or n = "ReviewCommentEvent") Then GoTo NextCheck
        If ver < 6 And (n = "OperationJournal" Or n = "NarrativeSnippet" Or n = "LibraryPreference" Or n = "FolderBatchChange") Then GoTo NextCheck
        If ver < 5 And n = "RecoveryDraft" Then GoTo NextCheck
        If ver < 4 And (n = "ImportReceipt" Or n = "ImportRawValue") Then GoTo NextCheck
        If ver = 1 And (n = "Folder" Or n = "FolderMove") Then GoTo NextCheck
        stage = "Verify " & n
        checkedRows = 0
        If showProgress Then Application.Run "WorkspaceImportProgress", "Verifying all fields in " & n & "..."
        sql = "SELECT * FROM [" & n & "] ORDER BY [" & s.TableDefs(n).Fields(0).Name & "]"
        If n = "PriorEvaluationTrait" Then sql = sql & ",BlockNumber"
        Set a = s.OpenRecordset(sql, dbOpenSnapshot): Set b = t.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not a.EOF
            If b.EOF Then Err.Raise vbObjectError + 304, , "Missing row in " & n
            For i = 0 To a.Fields.Count - 1
                av = a.Fields(i).Value: bv = b.Fields(a.Fields(i).Name).Value
                If IsNull(av) <> IsNull(bv) Then Err.Raise vbObjectError + 305, , "Null mismatch in " & n
                If Not IsNull(av) Then
                    If VarType(av) = vbString Then
                        If StrComp(CStr(av), CStr(bv), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 306, , "Text mismatch in " & n
                    Else
                        If av <> bv Then Err.Raise vbObjectError + 307, , "Value mismatch in " & n
                    End If
                End If
            Next
            a.MoveNext: b.MoveNext
            checkedRows = checkedRows + 1
            If showProgress And checkedRows Mod 100 = 0 Then Application.Run "WorkspaceImportProgress", "Verifying " & n & ": " & checkedRows & " records checked"
        Loop
        If Not b.EOF Then Err.Raise vbObjectError + 308, , "Extra rows in " & n
        a.Close: b.Close
NextCheck:
    Next
    If showProgress Then Application.Run "WorkspaceImportProgress", "Field verification finished in " & Format$(MigrationSeconds(verifyStarted), "0.00") & " seconds."
    If showProgress Then Application.Run "WorkspaceImportProgress", "Finalizing notes and member directory compatibility..."
    If targetVersion >= 22 Then Application.Eval "MigrateLegacyReportNotes()"
    If targetVersion >= 23 And ver < 23 Then Call SyncDirectoryFromReports()
    If targetVersion >= 26 Then Application.Eval "MigrateReportMessages()"
    If Len(expectedHash) > 0 Then
        If showProgress Then Application.Run "WorkspaceImportProgress", "All copied fields verified. Checking source is unchanged..."
        If SourceFileHash(sourcePath) <> expectedHash Then Err.Raise vbObjectError + 310, , "Source changed during transfer. No data imported."
    End If
    w.CommitTrans: trans = False
    RemoveMigrationQueries temporaryQuery
    s.Close: reader.Close
    MigrateWorkspace = "PASS: all original rows and fields verified; schema " & ver & " to " & targetVersion
    Exit Function
Failed:
    message = Err.Description
    Resume RollbackFailure
RollbackFailure:
    On Error Resume Next
    'Release readers/query handles before rollback invalidates their objects.
    If Not r Is Nothing Then r.Close
    If Not a Is Nothing Then a.Close
    If Not b Is Nothing Then b.Close
    If Not q Is Nothing Then q.Close
    If trans Then w.Rollback
    RemoveMigrationQueries temporaryQuery
    If Not s Is Nothing Then s.Close
    If Not reader Is Nothing Then reader.Close
    Err.Clear
    MigrateWorkspace = "FAIL: " & stage & ": " & message
End Function

Private Function IsMajorTable(ByVal name As String) As Boolean
    Dim key As Variant
    For Each key In Array("SummaryGroup", "GroupMember", "GroupSnapshot", "GroupSnapshotMember", "FieldDefault", "ChangeBatch", "ChangeItem", "RouteTemplate", "RouteTemplateStep", "RouteRun", "RouteRunStep", "ComparisonSet", "ComparisonMember", "MeetingProposal", "RulePack", "RuleDefinition", "ValidationRun", "ValidationFinding")
        If key = name Then IsMajorTable = True: Exit Function
    Next
End Function

Private Function IsRSTable(ByVal name As String) As Boolean
    IsRSTable = (name = "ReportingSeniorRecord" Or name = "ReportingSeniorValue" Or name = "RSCABaseline" Or name = "GroupRSCABaseline")
End Function

Private Function DeferredMigrationField(ByVal tableName As String, ByVal fieldName As String) As Boolean
    Select Case tableName & "." & fieldName
        Case "Report.CurrentRevisionID", "ReportRevision.ParentRevisionID", "Folder.ParentFolderID", "ReviewComment.ThreadID", "ReportMessage.ParentMessageID", "ReportNoteDraft.ParentMessageID"
            DeferredMigrationField = True
    End Select
End Function
Private Function MigrationSeconds(ByVal started As Single) As Double
    MigrationSeconds = CDbl(Timer) - CDbl(started)
    If MigrationSeconds < 0 Then MigrationSeconds = MigrationSeconds + 86400#
End Function

Private Sub RemoveMigrationQueries(ByVal prefix As String)
    Dim cleanup As DAO.Database, i As Long, queryName As String
    If Len(prefix) = 0 Then Exit Sub
    On Error Resume Next
    Set cleanup = CurrentDb
    For i = cleanup.QueryDefs.Count - 1 To 0 Step -1
        queryName = cleanup.QueryDefs(i).Name
        If Left$(queryName, Len(prefix)) = prefix Then cleanup.QueryDefs.Delete queryName
    Next
End Sub
