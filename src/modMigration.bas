Attribute VB_Name = "modMigration"
Option Compare Database
Option Explicit

Public Function MigrateWorkspace(ByVal sourcePath As String, Optional ByVal showProgress As Boolean = False, Optional ByVal replaceSetup As Boolean = False, Optional ByVal expectedHash As String = "") As String
    Dim s As DAO.Database, t As DAO.Database, w As DAO.Workspace, reader As DAO.Workspace
    Dim r As DAO.Recordset, q As DAO.QueryDef
    Dim tables As Variant, n As Variant, i As Long, ver As Long, sql As String
    Dim names As String, vals As String, trans As Boolean, message As String, stage As String, sourceSQL As String, temporaryQuery As String
    Dim fieldName As String, deferred As Variant, relationPair As Variant, copyStarted As Single, verifyStarted As Single
    Dim table As DAO.TableDef, known As Boolean, targetVersion As Long, copiedRows As Long, checkedRows As Long, committed As Boolean
    On Error GoTo Failed
    Set reader = DBEngine.CreateWorkspace("SwiftRead", "admin", "", dbUseJet)
    Set s = reader.OpenDatabase(sourcePath, False, True)
    Set t = CurrentDb: Set w = DBEngine.Workspaces(0)
    Set r = s.OpenRecordset("AppMetadata", dbOpenSnapshot): ver = r!SchemaVersion: r.Close
    If ver < 1 Or ver > 41 Then Err.Raise vbObjectError + 301, "modMigration.MigrateWorkspace", "Unsupported source schema."
    targetVersion = DLookup("SchemaVersion", "AppMetadata")
    If targetVersion < 21 Or targetVersion > 41 Then Err.Raise vbObjectError + 302, "modMigration.MigrateWorkspace", "Target schema must be 21 through 41."
    If ver > targetVersion Then Err.Raise vbObjectError + 302, "modMigration.MigrateWorkspace", "Migration cannot downgrade the source schema."
    tables = Array("UserProfile", "Workspace", "Folder", "Report", "ReportRevision", "RevisionValue", "RevisionChange", "ReportEvent", "FolderMove", "OperationJournal", "ImportReceipt", "ImportRawValue", "RecoveryDraft", "NarrativeSnippet", "LibraryPreference", "FolderBatchChange", "ReviewComment", "ReviewCommentEvent", "ReviewTask", "ReviewTaskEvent", _
        "SummaryGroup", "GroupMember", "GroupSnapshot", "GroupSnapshotMember", "FieldDefault", "ChangeBatch", "ChangeItem", "RouteTemplate", "RouteTemplateStep", "RouteRun", "RouteRunStep", "ComparisonSet", "ComparisonMember", "MeetingProposal", "RulePack", "RuleDefinition", "ValidationRun", "ValidationFinding", "ReportingSeniorRecord", "ReportingSeniorValue", "RSCABaseline", "GroupRSCABaseline")
    If targetVersion >= 22 Then tables = Split(Join(tables, "|") & "|ReportNote|ReportNoteDraft", "|")
    If targetVersion >= 23 Then tables = Split(Join(tables, "|") & "|CommandMember|MemberReport|PriorEvaluation|PriorEvaluationTrait", "|")
    If targetVersion >= 26 Then tables = Split(Join(tables, "|") & "|ReportMessage", "|")
    If targetVersion >= 27 Then tables = Split(Join(tables, "|") & "|DirectoryMerge|DirectoryMergeReport|DirectoryMergePrior", "|")
    If targetVersion >= 28 Then tables = Split(Join(tables, "|") & "|GroupCompletion|GroupCompletionReport|GroupCompletionCurrent|RSCountedReport|ReportHead|RevisionDate|RevisionChoice|FieldDefaultLink|GroupLifecycle", "|")
    If targetVersion >= 29 Then tables = Split(Join(tables, "|") & "|AuditActor|AuditEvent", "|")
    If targetVersion >= 30 Then tables = Split(Join(tables, "|") & "|WizardPreview", "|")
    If targetVersion >= 34 Then tables = Split(Join(tables, "|") & "|MemberCounseling|DirectoryMergeCounseling", "|")
    If targetVersion >= 35 Then tables = Split(Join(tables, "|") & "|ReportSignoff|DraftDisposal", "|")
    If targetVersion >= 36 Then tables = Split(Join(tables, "|") & "|DirectoryMergeSenior", "|")
    If targetVersion >= 37 Then tables = Split(Join(tables, "|") & "|RouteRunEvent", "|")
    If targetVersion >= 38 Then tables = Split(Join(tables, "|") & "|RosterImport|RosterPresence|MemberDeparture", "|")
    If targetVersion >= 39 Then tables = Split(Join(tables, "|") & "|ReportDueDismissal", "|")
    For Each table In s.TableDefs
        If Left$(table.Name, 4) <> "MSys" Then
            known = (table.name = "AppMetadata" Or (ver >= 31 And table.name = "TraitDefinition"))
            For Each n In tables
                If table.Name = n Then known = True
            Next
            If Not known Then Err.Raise vbObjectError + 309, "modMigration.MigrateWorkspace", "Unrecognized source table: " & table.name & ". Migration stopped to prevent data loss."
        End If
    Next
    'One inexpensive guard; TOP 1 never counts all records in a populated table.
    For Each table In t.TableDefs
        If left$(table.name, 4) <> "MSys" And table.name <> "AppMetadata" And table.name <> "TraitDefinition" Then
            known = False
            If replaceSetup And targetVersion >= 30 And table.Name = "WizardPreview" Then known = WizardImportStagingOnly()
            If Not (replaceSetup And (table.Name = "Workspace" Or table.Name = "UserProfile")) And Not known Then
                Set r = t.OpenRecordset("SELECT TOP 1 * FROM [" & table.Name & "]", dbOpenSnapshot)
                If Not r.EOF Then r.Close: Err.Raise vbObjectError + 303, "modMigration.MigrateWorkspace", "Target must be blank."
                r.Close
            End If
        End If
    Next
    'Saved query metadata must be created outside the data transaction. Otherwise
    'an action-query failure can invalidate DAO handles while rolling back DDL.
    deferred = Array("Report|CurrentRevisionID|ReportID", "ReportRevision|ParentRevisionID|RevisionID", "Folder|ParentFolderID|FolderID", "ReviewComment|ThreadID|CommentID")
    If ver >= 26 Then deferred = Split(Join(deferred, "~") & "~ReportMessage|ParentMessageID|MessageID~ReportNoteDraft|ParentMessageID|DraftGuid", "~")
    If ver >= 28 Then deferred = Split(Join(deferred, "~") & "~RSCABaseline|ParentBaselineID|BaselineID~GroupCompletion|PriorCompletionID|CompletionID", "~")
    'The senior's member link waits until CommandMember is copied.
    If ver >= 36 Then deferred = Split(Join(deferred, "~") & "~ReportingSeniorRecord|MemberID|SeniorID", "~")
    temporaryQuery = "zSwiftMigration" & Left$(Replace(Replace(Replace(NewIdentity(), "{", ""), "}", ""), "-", ""), 24)
    sourceSQL = " IN " & Chr$(34) & sourcePath & Chr$(34)
    'Prepare read-only external-source queries before the data transaction.
    'The same queries support relationship restoration and engine verification.
    For Each table In s.TableDefs
        If Left$(table.Name, 4) <> "MSys" And table.Name <> "AppMetadata" Then
            Set q = t.CreateQueryDef(temporaryQuery & table.Name, "SELECT * FROM [" & table.Name & "]" & sourceSQL)
            q.Close
        End If
    Next
    w.BeginTrans: trans = True
    If ver >= 31 Then
        stage = "Preserve trait definitions"
        ImportTraitDefinitions s, t
    End If
    'Only the current import descriptor may occupy an otherwise blank target.
    'Its removal rolls back with the import; all source preview rows are copied.
    If replaceSetup And targetVersion >= 30 Then t.Execute "DELETE FROM WizardPreview", dbFailOnError
    If replaceSetup Then
        t.Execute "DELETE FROM Workspace", dbFailOnError
        t.Execute "DELETE FROM UserProfile", dbFailOnError
    End If
    copyStarted = Timer
    For Each n In tables
        If ver < 39 And n = "ReportDueDismissal" Then GoTo NextTable
        If ver < 38 And (n = "RosterImport" Or n = "RosterPresence" Or n = "MemberDeparture") Then GoTo NextTable
        If ver < 37 And n = "RouteRunEvent" Then GoTo NextTable
        If ver < 36 And n = "DirectoryMergeSenior" Then GoTo NextTable
        If ver < 35 And (n = "ReportSignoff" Or n = "DraftDisposal") Then GoTo NextTable
        If ver < 34 And (n = "MemberCounseling" Or n = "DirectoryMergeCounseling") Then GoTo NextTable
        If ver < 30 And n = "WizardPreview" Then GoTo NextTable
        If ver < 29 And (n = "AuditActor" Or n = "AuditEvent") Then GoTo NextTable
        If ver < 28 And IsPersistenceTable(CStr(n)) Then GoTo NextTable
        If ver < 27 And Left$(CStr(n), 14) = "DirectoryMerge" Then GoTo NextTable
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
        If targetVersion >= 28 And ver < 28 And n = "Report" Then
            names = names & ",[LegacyFormKind],[LegacyWorkflowState],[LegacyWorkflowNull]"
            vals = vals & ",[FormKind],[WorkflowState],IIf([WorkflowState] Is Null,True,False)"
        End If
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
        If ver < 39 And n = "ReportDueDismissal" Then GoTo NextCheck
        If ver < 38 And (n = "RosterImport" Or n = "RosterPresence" Or n = "MemberDeparture") Then GoTo NextCheck
        If ver < 37 And n = "RouteRunEvent" Then GoTo NextCheck
        If ver < 36 And n = "DirectoryMergeSenior" Then GoTo NextCheck
        If ver < 35 And (n = "ReportSignoff" Or n = "DraftDisposal") Then GoTo NextCheck
        If ver < 34 And (n = "MemberCounseling" Or n = "DirectoryMergeCounseling") Then GoTo NextCheck
        If ver < 30 And n = "WizardPreview" Then GoTo NextCheck
        If ver < 29 And (n = "AuditActor" Or n = "AuditEvent") Then GoTo NextCheck
        If ver < 28 And IsPersistenceTable(CStr(n)) Then GoTo NextCheck
        If ver < 27 And Left$(CStr(n), 14) = "DirectoryMerge" Then GoTo NextCheck
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
        checkedRows = VerifyMigrationTable(t, s.TableDefs(CStr(n)), temporaryQuery & CStr(n), showProgress)
        If showProgress Then Application.Run "WorkspaceImportProgress", "Verified " & n & ": " & checkedRows & " records checked by database engine"
NextCheck:
    Next
    If showProgress Then Application.Run "WorkspaceImportProgress", "Field verification finished in " & Format$(MigrationSeconds(verifyStarted), "0.00") & " seconds."
    If showProgress Then Application.Run "WorkspaceImportProgress", "Finalizing notes and member directory compatibility..."
    If targetVersion >= 28 Then
        MigrateRSAccounting
        MigrateReportState
    End If
    'Direct calls: an error raised through Application.Eval never reaches Failed
    'and shows an unhandled run-time dialog instead of rolling back the import.
    If targetVersion >= 22 And ver < 29 Then Call MigrateLegacyReportNotes
    If targetVersion >= 23 And ver < 23 Then Call MigrateDirectoryCompatibility()
    If targetVersion >= 26 Then Call MigrateReportMessages
    If targetVersion >= 29 Then MigrateAuditHistory
    If Len(expectedHash) > 0 Then
        If showProgress Then Application.Run "WorkspaceImportProgress", "All copied fields verified. Checking source is unchanged..."
        If SourceFileHash(sourcePath) <> expectedHash Then Err.Raise vbObjectError + 310, "modMigration.MigrateWorkspace", "Source changed during transfer. No data imported."
    End If
    w.CommitTrans: trans = False: committed = True
    RemoveMigrationQueries temporaryQuery
    s.Close: reader.Close
    ResetActiveContributor
    UXLoadContext
    MigrateWorkspace = "PASS: all original rows and fields verified; schema " & ver & " to " & targetVersion
    Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modMigration", "MigrateWorkspace", stage, Err.Number, Err.Description, Err.Source, IIf(committed, "Committed", IIf(trans, "Active", "Not started")))
    message = CStr(swiftPrimary("Description"))
    Resume RollbackFailure
RollbackFailure:
    On Error Resume Next
    'Release readers/query handles before rollback invalidates their objects.
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    If Not q Is Nothing Then SwiftClose q, swiftPrimary, "q"
    If trans Then SwiftRollback w, swiftPrimary
    RemoveMigrationQueries temporaryQuery
    If Not s Is Nothing Then SwiftClose s, swiftPrimary, "s"
    If Not reader Is Nothing Then SwiftClose reader, swiftPrimary, "reader"
    Err.Clear
    On Error GoTo 0
    MigrateWorkspace = "FAIL: " & SwiftFailureText(swiftPrimary)
End Function

 'Exact set verification: indexed joins and binary string comparisons run in ACE.
'VBA builds a bounded number of queries and reads only counts/mismatch results.
'Chunk numeric primary keys so the UI gets a guarded event turn between batches.
Public Function VerifyMigrationTable(ByVal db As DAO.Database, ByVal sourceTable As DAO.TableDef, ByVal sourceQuery As String, Optional ByVal showProgress As Boolean = False) As Long
    Dim idx As DAO.Index, key As DAO.Field, fld As DAO.Field, rows As DAO.Recordset
    Dim keys As Collection, join As String, firstKey As String, predicate As String
    Dim sourceCount As Long, targetCount As Long, low As Long, high As Long, haveLow As Boolean
    Dim numericKey As Boolean, range As String, boundary As String, fromSQL As String
    Dim first As Long, last As Long, i As Long, part As String, checked As Long, chunkCount As Long
    Dim failure As Long, message As String
    Const BatchRows As Long = 8192
    Const BatchFields As Long = 12
    On Error GoTo Failed
    Set keys = New Collection
    For Each idx In sourceTable.Indexes
        If idx.Primary Then
            For Each key In idx.Fields
                keys.Add key.Name
            Next
            Exit For
        End If
    Next
    If keys.count = 0 Then Err.Raise vbObjectError + 311, "modMigration.VerifyMigrationTable", "Source table has no primary key: " & sourceTable.name
    firstKey = CStr(keys(1))
    For i = 1 To keys.count
        If Len(join) > 0 Then join = join & " AND "
        join = join & "src.[" & keys(i) & "]=dst.[" & keys(i) & "]"
    Next
    Set rows = db.OpenRecordset("SELECT Count(*) AS N FROM [" & sourceQuery & "]", dbOpenSnapshot)
    sourceCount = rows!n: rows.Close
    Set rows = db.OpenRecordset("SELECT Count(*) AS N FROM [" & sourceTable.Name & "]", dbOpenSnapshot)
    targetCount = rows!n: rows.Close
    If sourceCount <> targetCount Then Err.Raise vbObjectError + 304, "modMigration.VerifyMigrationTable", "Row count mismatch in " & sourceTable.name
    If sourceCount = 0 Then Exit Function
    numericKey = (sourceTable.Fields(firstKey).Type = dbLong)
    fromSQL = " FROM [" & sourceQuery & "] AS src LEFT JOIN [" & sourceTable.Name & "] AS dst ON " & join
    Do
        range = "": chunkCount = sourceCount
        If numericKey Then
            boundary = ""
            If haveLow Then boundary = " WHERE [" & firstKey & "]>" & low
            Set rows = db.OpenRecordset("SELECT TOP " & BatchRows & " [" & firstKey & "] FROM [" & sourceQuery & "]" & boundary & " ORDER BY [" & firstKey & "]", dbOpenSnapshot)
            If rows.EOF Then rows.Close: Exit Do
            rows.MoveLast: high = rows.Fields(0): rows.MoveFirst
            'The boundary query returns at most BatchRows keys, never field values.
            range = "src.[" & firstKey & "]<=" & high & " AND "
            If haveLow Then range = "src.[" & firstKey & "]>" & low & " AND " & range
            chunkCount = rows.RecordCount: rows.Close
        End If
        For first = 0 To sourceTable.Fields.count - 1 Step BatchFields
            last = first + BatchFields - 1
            If last >= sourceTable.Fields.count Then last = sourceTable.Fields.count - 1
            predicate = "dst.[" & firstKey & "] Is Null"
            For i = first To last
                Set fld = sourceTable.Fields(i)
                part = MigrationDifference(fld)
                predicate = predicate & " OR " & part
            Next
            Set rows = db.OpenRecordset("SELECT TOP 1 src.[" & firstKey & "]" & fromSQL & " WHERE " & range & "(" & predicate & ")", dbOpenSnapshot)
            If Not rows.EOF Then Err.Raise vbObjectError + 306, "modMigration.VerifyMigrationTable", "Missing row or field mismatch in " & sourceTable.name & " (fields " & first + 1 & "-" & last + 1 & ")."
            rows.Close
            If showProgress Then Application.Run "WorkspaceImportPulse"
        Next
        checked = checked + chunkCount
        If showProgress Then Application.Run "WorkspaceImportProgress", "Verifying " & sourceTable.Name & ": " & checked & " records checked by database engine"
        If Not numericKey Then Exit Do
        low = high: haveLow = True
    Loop
    VerifyMigrationTable = sourceCount
    Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    Resume CleanupVerification
CleanupVerification:
    On Error Resume Next
    If Not rows Is Nothing Then rows.Close
    On Error GoTo 0
    Err.Raise failure, "VerifyMigrationTable", message
End Function
Private Function MigrationDifference(ByVal field As DAO.Field) As String
    Dim a As String, b As String, compare As String
    a = "src.[" & field.Name & "]": b = "dst.[" & field.Name & "]"
    Select Case field.Type
        Case dbText, dbMemo
            'Binary comparison preserves case, Unicode, trailing spaces and full memos.
            compare = "StrComp(" & a & "," & b & ",0)<>0"
        Case dbBoolean, dbByte, dbInteger, dbLong, dbCurrency, dbSingle, dbDouble, dbDate, dbDecimal, dbGUID
            compare = a & "<>" & b
        Case Else
            Err.Raise vbObjectError + 312, "modMigration.MigrationDifference", "Unsupported verification field type: " & field.name & " (" & field.Type & ")"
    End Select
    MigrationDifference = "((" & a & " Is Null)<>(" & b & " Is Null) OR " & compare & ")"
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

Private Function IsPersistenceTable(ByVal name As String) As Boolean
    Dim key As Variant
    For Each key In Array("GroupCompletion", "GroupCompletionReport", "GroupCompletionCurrent", "RSCountedReport", "ReportHead", "RevisionDate", "RevisionChoice", "FieldDefaultLink", "GroupLifecycle")
        If key = name Then IsPersistenceTable = True: Exit Function
    Next
End Function

Private Function DeferredMigrationField(ByVal tableName As String, ByVal fieldName As String) As Boolean
    Select Case tableName & "." & fieldName
        Case "Report.CurrentRevisionID", "ReportRevision.ParentRevisionID", "Folder.ParentFolderID", "ReviewComment.ThreadID", "ReportMessage.ParentMessageID", "ReportNoteDraft.ParentMessageID", "RSCABaseline.ParentBaselineID", "GroupCompletion.PriorCompletionID", "ReportingSeniorRecord.MemberID"
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
