Attribute VB_Name = "modMigration"
Option Compare Database
Option Explicit

Public Function MigrateWorkspace(ByVal sourcePath As String, Optional ByVal showProgress As Boolean = False, Optional ByVal replaceSetup As Boolean = False, Optional ByVal expectedHash As String = "") As String
    Dim s As DAO.Database, t As DAO.Database, w As DAO.Workspace, reader As DAO.Workspace
    Dim r As DAO.Recordset, q As DAO.QueryDef
    Dim tables As Variant, n As Variant, i As Long, ver As Long, sql As String
    Dim names As String, vals As String, trans As Boolean, message As String, stage As String, sourceSQL As String, temporaryQuery As String
    Dim fieldName As String, deferred As Variant, relationPair As Variant, copyStarted As Single, verifyStarted As Single
    Dim table As DAO.TableDef, known As Boolean, targetVersion As Long, copiedRows As Long, checkedRows As Long, committed As Boolean, restored As Long
    On Error GoTo Failed
    Set reader = DBEngine.CreateWorkspace("SwiftRead", "admin", "", dbUseJet)
    Set s = reader.OpenDatabase(sourcePath, False, True)
    Set t = CurrentDb: Set w = DBEngine.Workspaces(0)
    Set r = s.OpenRecordset("AppMetadata", dbOpenSnapshot): ver = r!SchemaVersion: r.Close
    'Schema 35 (release 0.49.0, the first public release) is the oldest imported.
    'Release 0.55.0 still imports schemas 1-48 and bridges older workspaces. The source
    'is at least 35 and never newer than the target, so no older-schema branches remain.
    If ver >= 1 And ver < 35 Then Err.Raise vbObjectError + 301, "modMigration.MigrateWorkspace", "Unsupported source schema " & ver & ". This version imports workspaces from release 0.49.0 (schema 35) onward. Import the file into a Swift NAVFIT 0.55.0 Blank first, then import that workspace here."
    If ver < 35 Or ver > 50 Then Err.Raise vbObjectError + 301, "modMigration.MigrateWorkspace", "Unsupported source schema."
    targetVersion = DLookup("SchemaVersion", "AppMetadata")
    'The target is always this version's own workspace, so its bound follows SCHEMA_VERSION.
    If targetVersion < 35 Or targetVersion > SCHEMA_VERSION Then Err.Raise vbObjectError + 302, "modMigration.MigrateWorkspace", "Target schema must be 35 through " & SCHEMA_VERSION & "."
    If ver > targetVersion Then Err.Raise vbObjectError + 302, "modMigration.MigrateWorkspace", "Migration cannot downgrade the source schema."
    tables = Array("UserProfile", "Workspace", "Folder", "Report", "ReportRevision", "RevisionValue", "RevisionChange", "ReportEvent", "FolderMove", "OperationJournal", "ImportReceipt", "ImportRawValue", "RecoveryDraft", "NarrativeSnippet", "LibraryPreference", "FolderBatchChange", "ReviewComment", "ReviewCommentEvent", "ReviewTask", "ReviewTaskEvent", _
        "SummaryGroup", "GroupMember", "GroupSnapshot", "GroupSnapshotMember", "FieldDefault", "ChangeBatch", "ChangeItem", "RouteTemplate", "RouteTemplateStep", "RouteRun", "RouteRunStep", "ComparisonSet", "ComparisonMember", "MeetingProposal", "RulePack", "RuleDefinition", "ValidationRun", "ValidationFinding", "ReportingSeniorRecord", "ReportingSeniorValue", "RSCABaseline", "GroupRSCABaseline", _
        "ReportNote", "ReportNoteDraft", "CommandMember", "MemberReport", "PriorEvaluation", "PriorEvaluationTrait", "ReportMessage", "DirectoryMerge", "DirectoryMergeReport", "DirectoryMergePrior", "GroupCompletion", "GroupCompletionReport", "GroupCompletionCurrent", "RSCountedReport", "ReportHead", "RevisionDate", "RevisionChoice", "FieldDefaultLink", "GroupLifecycle", "AuditActor", "AuditEvent", "WizardPreview", "MemberCounseling", "DirectoryMergeCounseling", "ReportSignoff", "DraftDisposal")
    If targetVersion >= 36 Then tables = Split(Join(tables, "|") & "|DirectoryMergeSenior", "|")
    If targetVersion >= 37 Then tables = Split(Join(tables, "|") & "|RouteRunEvent", "|")
    If targetVersion >= 38 Then tables = Split(Join(tables, "|") & "|RosterImport|RosterPresence|MemberDeparture", "|")
    If targetVersion >= 39 Then tables = Split(Join(tables, "|") & "|ReportDueDismissal", "|")
    'Groups can reference an ISIC, so copy that parent before SummaryGroup.
    If targetVersion >= 42 Then tables = Split(Replace(Join(tables, "|"), "|SummaryGroup|", "|ISIC|SummaryGroup|"), "|")
    If targetVersion >= 43 Then tables = Split(Join(tables, "|") & "|DirectoryMergeLink", "|")
    If targetVersion >= 44 Then tables = Split(Join(tables, "|") & "|SavedReportFilter", "|")
    If targetVersion >= 45 Then tables = Split(Join(tables, "|") & "|RevisionRetention", "|")
    If targetVersion >= 47 Then tables = Split(Join(tables, "|") & "|RestrictedException|RestrictedDismissal", "|")
    If targetVersion >= 48 Then tables = Split(Join(tables, "|") & "|RestrictedCheckOff", "|")
    For Each table In s.TableDefs
        If Left$(table.Name, 4) <> "MSys" Then
            'Who had the source open is momentary; it is known but never copied.
            known = (table.name = "AppMetadata" Or table.name = "TraitDefinition" Or SwiftTransientTable(table.Name))
            For Each n In tables
                If table.Name = n Then known = True
            Next
            If Not known Then Err.Raise vbObjectError + 309, "modMigration.MigrateWorkspace", "Unrecognized source table: " & table.name & ". Migration stopped to prevent data loss."
        End If
    Next
    'One inexpensive guard; TOP 1 never counts all records in a populated table.
    For Each table In t.TableDefs
        If SwiftWorkspaceTable(table.Name) And Not SwiftTransientTable(table.Name) And table.name <> "AppMetadata" And table.name <> "TraitDefinition" Then
            known = False
            If replaceSetup And table.Name = "WizardPreview" Then known = WizardImportStagingOnly()
            If Not (replaceSetup And (table.Name = "Workspace" Or table.Name = "UserProfile")) And Not known Then
                Set r = t.OpenRecordset("SELECT TOP 1 * FROM [" & table.Name & "]", dbOpenSnapshot)
                If Not r.EOF Then r.Close: Err.Raise vbObjectError + 303, "modMigration.MigrateWorkspace", "Target must be blank."
                r.Close
            End If
        End If
    Next
    'Saved query metadata must be created outside the data transaction. Otherwise
    'an action-query failure can invalidate DAO handles while rolling back DDL.
    deferred = Array("Report|CurrentRevisionID|ReportID", "ReportRevision|ParentRevisionID|RevisionID", "Folder|ParentFolderID|FolderID", "ReviewComment|ThreadID|CommentID", _
        "ReportMessage|ParentMessageID|MessageID", "ReportNoteDraft|ParentMessageID|DraftGuid", "RSCABaseline|ParentBaselineID|BaselineID", "GroupCompletion|PriorCompletionID|CompletionID")
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
    stage = "Preserve trait definitions"
    ImportTraitDefinitions s, t
    'Only the current import descriptor may occupy an otherwise blank target.
    'Its removal rolls back with the import; all source preview rows are copied.
    If replaceSetup Then t.Execute "DELETE FROM WizardPreview", dbFailOnError
    If replaceSetup Then
        t.Execute "DELETE FROM Workspace", dbFailOnError
        t.Execute "DELETE FROM UserProfile", dbFailOnError
    End If
    copyStarted = Timer
    For Each n In tables
        If ver < 48 And n = "RestrictedCheckOff" Then GoTo NextTable
        If ver < 47 And (n = "RestrictedException" Or n = "RestrictedDismissal") Then GoTo NextTable
        If ver < 45 And n = "RevisionRetention" Then GoTo NextTable
        If ver < 44 And n = "SavedReportFilter" Then GoTo NextTable
        If ver < 43 And n = "DirectoryMergeLink" Then GoTo NextTable
        If ver < 42 And n = "ISIC" Then GoTo NextTable
        If ver < 39 And n = "ReportDueDismissal" Then GoTo NextTable
        If ver < 38 And (n = "RosterImport" Or n = "RosterPresence" Or n = "MemberDeparture") Then GoTo NextTable
        If ver < 37 And n = "RouteRunEvent" Then GoTo NextTable
        If ver < 36 And n = "DirectoryMergeSenior" Then GoTo NextTable
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
        stage = "Restore " & relationPair(0) & "." & relationPair(1)
        t.Execute "UPDATE [" & relationPair(0) & "] AS dst INNER JOIN [" & temporaryQuery & relationPair(0) & "] AS src ON dst.[" & relationPair(2) & "]=src.[" & relationPair(2) & "] SET dst.[" & relationPair(1) & "]=src.[" & relationPair(1) & "]", dbFailOnError
    Next
    If showProgress Then Application.Run "WorkspaceImportProgress", "Bulk copying finished in " & Format$(MigrationSeconds(copyStarted), "0.00") & " seconds. Verifying copied values..."
    verifyStarted = Timer
    For Each n In tables
        If ver < 48 And n = "RestrictedCheckOff" Then GoTo NextCheck
        If ver < 47 And (n = "RestrictedException" Or n = "RestrictedDismissal") Then GoTo NextCheck
        If ver < 45 And n = "RevisionRetention" Then GoTo NextCheck
        If ver < 44 And n = "SavedReportFilter" Then GoTo NextCheck
        If ver < 43 And n = "DirectoryMergeLink" Then GoTo NextCheck
        If ver < 42 And n = "ISIC" Then GoTo NextCheck
        If ver < 39 And n = "ReportDueDismissal" Then GoTo NextCheck
        If ver < 38 And (n = "RosterImport" Or n = "RosterPresence" Or n = "MemberDeparture") Then GoTo NextCheck
        If ver < 37 And n = "RouteRunEvent" Then GoTo NextCheck
        If ver < 36 And n = "DirectoryMergeSenior" Then GoTo NextCheck
        stage = "Verify " & n
        checkedRows = 0
        If showProgress Then Application.Run "WorkspaceImportProgress", "Verifying all fields in " & n & "..."
        checkedRows = VerifyMigrationTable(t, s.TableDefs(CStr(n)), temporaryQuery & CStr(n), showProgress)
        If showProgress Then Application.Run "WorkspaceImportProgress", "Verified " & n & ": " & checkedRows & " records checked by database engine"
NextCheck:
    Next
    If showProgress Then Application.Run "WorkspaceImportProgress", "Field verification finished in " & Format$(MigrationSeconds(verifyStarted), "0.00") & " seconds."
    If showProgress Then Application.Run "WorkspaceImportProgress", "Finalizing notes and member directory compatibility..."
    MigrateRSAccounting
    MigrateReportState
    'Direct calls: an error raised through Application.Eval never reaches Failed
    'and shows an unhandled run-time dialog instead of rolling back the import.
    Call MigrateReportMessages
    'Schema 46 adds EVAL block 44. Before audit history so its events are audited.
    If targetVersion >= 46 And ver < 46 Then
        stage = "Restore block 44"
        If showProgress Then Application.Run "WorkspaceImportProgress", "Restoring EVAL block 44 from NAVFIT98A imports..."
        restored = MigrateQualifications()
    End If
    MigrateAuditHistory
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
    If restored > 0 Then MigrateWorkspace = MigrateWorkspace & "; block 44 restored on " & restored & IIf(restored = 1, " report", " reports") & " from NAVFIT98A imports"
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

'Schema 46 adds EVAL block 44 (Qualifications). Each NAVFIT98A full import kept the
'native Qualifications column in ImportRawValue (read with NativeReportMemo). For an EVAL whose latest active full
'import carried text, add that text to the import's revision and each later one that
'does not record the field, then to the report cache when its head is among them.
'Disposed revisions, earlier revisions and recorded values are left alone. A report
'event records the restore. Returns the number of reports restored.
Public Function MigrateQualifications() As Long
    Dim db As DAO.Database, reports As DAO.Recordset, raw As DAO.Recordset, revisions As DAO.Recordset, values As DAO.Recordset, e As DAO.Recordset
    Dim receipt As Variant, text As String, head As Long, headRestored As Boolean, author As Long, applied As Long
    Set db = CurrentDb
    Set reports = db.OpenRecordset("SELECT ReportID FROM Report WHERE FormKind='EVAL' ORDER BY ReportID", dbOpenSnapshot)
    Do While Not reports.EOF
        receipt = MajorScalar("SELECT Max(ImportReceiptID) FROM ImportReceipt WHERE ReportID=" & reports!ReportID & " AND IsUndone=False AND AdapterVersion='navfit-full-v1/local-CStr'")
        If Not IsNull(receipt) Then
            text = "": applied = 0
            text = Nz(NativeReportMemo(CLng(receipt), "Qualifications"), "")
            Set raw = db.OpenRecordset("SELECT AppliedRevisionID,AuthorProfileID FROM ImportReceipt WHERE ImportReceiptID=" & CLng(receipt), dbOpenSnapshot)
            applied = Nz(raw!AppliedRevisionID, 0): author = raw!AuthorProfileID
            raw.Close
            If Len(Trim$(text)) > 0 And applied > 0 Then
                head = ReportHeadRevision(reports!ReportID): headRestored = False
                Set revisions = db.OpenRecordset("SELECT RevisionID FROM ReportRevision WHERE ReportID=" & reports!ReportID & " AND RevisionID>=" & applied & " AND Not Nz(ValuesDisposed,False) AND (Nz(ValueEncoding,0)=0 Or Nz(FieldDefinitionVersion,0)<3) AND RevisionID Not In (SELECT RevisionID FROM RevisionValue WHERE FieldKey='Qualifications') ORDER BY RevisionID", dbOpenSnapshot)
                If Not revisions.EOF Then
                    Set values = db.OpenRecordset("RevisionValue", dbOpenDynaset, dbAppendOnly)
                    Do While Not revisions.EOF
                        values.AddNew: values!RevisionID = revisions!RevisionID: values!FieldKey = "Qualifications": values!ValueText = text: values!IsNullValue = False: values.Update
                        If Not IsNull(MajorScalar("SELECT ValueHash FROM ReportRevision WHERE RevisionID=" & revisions!RevisionID)) Then SealRevision db, CLng(revisions!RevisionID)
                        If revisions!RevisionID = head Then headRestored = True
                        revisions.MoveNext
                    Loop
                    values.Close
                    If headRestored Then db.Execute "UPDATE Report SET Qualifications=" & QuoteSQL(text) & " WHERE ReportID=" & reports!ReportID, dbFailOnError
                    Set e = db.OpenRecordset("ReportEvent", dbOpenDynaset)
                    e.AddNew: e!ReportID = reports!ReportID: e!AuthorProfileID = author: e!EventAt = Now(): e!EventKind = "Block 44 restored from NAVFIT98A": e.Update: e.Close
                    MigrateQualifications = MigrateQualifications + 1
                End If
                revisions.Close
            End If
        End If
        reports.MoveNext
    Loop
    reports.Close
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
