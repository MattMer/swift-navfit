Attribute VB_Name = "modMigration"
Option Compare Database
Option Explicit

Private Function SqlType(ByVal f As DAO.Field) As String
    Select Case f.Type
        Case dbBoolean: SqlType = "Bit"
        Case dbLong: SqlType = "Long"
        Case dbInteger: SqlType = "Short"
        Case dbDouble: SqlType = "Double"
        Case dbDate: SqlType = "DateTime"
        Case dbText: SqlType = "Text(" & f.Size & ")"
        Case dbMemo: SqlType = "LongText"
        Case Else: Err.Raise vbObjectError + 300, , "Unsupported field type."
    End Select
End Function

Public Function MigrateWorkspace(ByVal sourcePath As String) As String
    Dim s As DAO.Database, t As DAO.Database, w As DAO.Workspace, reader As DAO.Workspace
    Dim r As DAO.Recordset, a As DAO.Recordset, b As DAO.Recordset, q As DAO.QueryDef
    Dim tables As Variant, n As Variant, i As Long, ver As Long, sql As String
    Dim names As String, defs As String, vals As String, trans As Boolean, message As String, stage As String, rowWhere As String, orderFields As String
    Dim av As Variant, bv As Variant, requiredFields As Collection, fieldIndex As Variant
    Dim table As DAO.TableDef, known As Boolean, targetVersion As Long
    On Error GoTo Failed
    Set reader = DBEngine.CreateWorkspace("SwiftRead", "admin", "", dbUseJet)
    Set s = reader.OpenDatabase(sourcePath, False, True)
    Set t = CurrentDb: Set w = DBEngine.Workspaces(0)
    Set r = s.OpenRecordset("AppMetadata", dbOpenSnapshot): ver = r!SchemaVersion: r.Close
    If ver < 1 Or ver > 25 Then Err.Raise vbObjectError + 301, , "Unsupported source schema."
    targetVersion = DLookup("SchemaVersion", "AppMetadata")
    If targetVersion <> 21 And targetVersion <> 22 And targetVersion <> 23 And targetVersion <> 24 And targetVersion <> 25 Then Err.Raise vbObjectError + 302, , "Target schema must be 21, 22, 23, 24 or 25."
    If ver > targetVersion Then Err.Raise vbObjectError + 302, , "Migration cannot downgrade the source schema."
    tables = Array("UserProfile", "Workspace", "Folder", "Report", "ReportRevision", "RevisionValue", "RevisionChange", "ReportEvent", "FolderMove", "OperationJournal", "ImportReceipt", "ImportRawValue", "RecoveryDraft", "NarrativeSnippet", "LibraryPreference", "FolderBatchChange", "ReviewComment", "ReviewCommentEvent", "ReviewTask", "ReviewTaskEvent", _
        "SummaryGroup", "GroupMember", "GroupSnapshot", "GroupSnapshotMember", "FieldDefault", "ChangeBatch", "ChangeItem", "RouteTemplate", "RouteTemplateStep", "RouteRun", "RouteRunStep", "ComparisonSet", "ComparisonMember", "MeetingProposal", "RulePack", "RuleDefinition", "ValidationRun", "ValidationFinding", "ReportingSeniorRecord", "ReportingSeniorValue", "RSCABaseline", "GroupRSCABaseline")
    If targetVersion >= 22 Then tables = Split(Join(tables, "|") & "|ReportNote|ReportNoteDraft", "|")
    If targetVersion >= 23 Then tables = Split(Join(tables, "|") & "|CommandMember|MemberReport|PriorEvaluation|PriorEvaluationTrait", "|")
    For Each table In s.TableDefs
        If Left$(table.Name, 4) <> "MSys" Then
            known = (table.Name = "AppMetadata")
            For Each n In tables
                If table.Name = n Then known = True
            Next
            If Not known Then Err.Raise vbObjectError + 309, , "Unrecognized source table: " & table.Name & ". Migration stopped to prevent data loss."
        End If
    Next
    For Each n In tables
        If DCount("*", CStr(n)) <> 0 Then Err.Raise vbObjectError + 303, , "Target must be blank."
    Next
    w.BeginTrans: trans = True
    For Each n In tables
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
        stage = "Copy " & n
        Set r = s.OpenRecordset("SELECT * FROM [" & n & "] ORDER BY [" & s.TableDefs(n).Fields(0).Name & "]", dbOpenSnapshot)
        names = "": defs = "": vals = "": Set requiredFields = New Collection
        For i = 0 To r.Fields.Count - 1
            If i = 0 Or r.Fields(i).Required Then
                If Len(names) > 0 Then names = names & ",": defs = defs & ",": vals = vals & ","
                names = names & "[" & r.Fields(i).Name & "]"
                defs = defs & "p" & i & " " & SqlType(r.Fields(i)): vals = vals & "p" & i
                requiredFields.Add i
            End If
        Next
        Do While Not r.EOF
            Set q = t.CreateQueryDef("", "PARAMETERS " & defs & "; INSERT INTO [" & n & "] (" & names & ") VALUES (" & vals & ")")
            For Each fieldIndex In requiredFields
                q.Parameters("p" & fieldIndex) = r.Fields(fieldIndex).Value
            Next
            stage = "Insert " & n & " record " & r.Fields(0).Value
            q.Execute dbFailOnError: q.Close
            rowWhere = "[" & r.Fields(0).Name & "]=" & r.Fields(0).Value
            If n = "PriorEvaluationTrait" Then rowWhere = rowWhere & " AND BlockNumber=" & r!BlockNumber
            Set b = t.OpenRecordset("SELECT * FROM [" & n & "] WHERE " & rowWhere, dbOpenDynaset)
            b.Edit
            For i = 1 To r.Fields.Count - 1
                stage = "Copy " & n & "." & r.Fields(i).Name
                If Not ((n = "Report" And r.Fields(i).Name = "CurrentRevisionID") Or (n = "Folder" And r.Fields(i).Name = "ParentFolderID")) Then b.Fields(r.Fields(i).Name) = r.Fields(i).Value
            Next
            b.Update: b.Close: r.MoveNext
        Loop
        r.Close
NextTable:
    Next
    Set r = s.OpenRecordset("SELECT ReportID,CurrentRevisionID FROM Report", dbOpenSnapshot)
    Set q = t.CreateQueryDef("", "PARAMETERS rev Long, id Long; UPDATE Report SET CurrentRevisionID=rev WHERE ReportID=id")
    Do While Not r.EOF
        q!rev = r!CurrentRevisionID: q!id = r!ReportID: q.Execute dbFailOnError: r.MoveNext
    Loop
    q.Close: r.Close
    If ver >= 2 Then
        Set r = s.OpenRecordset("SELECT FolderID,ParentFolderID FROM Folder", dbOpenSnapshot)
        Set q = t.CreateQueryDef("", "PARAMETERS parent Long, id Long; UPDATE Folder SET ParentFolderID=parent WHERE FolderID=id")
        Do While Not r.EOF
            q!parent = r!ParentFolderID: q!id = r!FolderID: q.Execute dbFailOnError: r.MoveNext
        Loop
        q.Close: r.Close
    End If
    For Each n In tables
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
        Loop
        If Not b.EOF Then Err.Raise vbObjectError + 308, , "Extra rows in " & n
        a.Close: b.Close
NextCheck:
    Next
    If targetVersion >= 22 Then Application.Eval "MigrateLegacyReportNotes()"
    If targetVersion >= 23 And ver < 23 Then Call SyncDirectoryFromReports()
    w.CommitTrans: trans = False: s.Close: reader.Close
    MigrateWorkspace = "PASS: all original rows and fields verified; schema " & ver & " to " & targetVersion
    Exit Function
Failed:
    message = Err.Description
    If trans Then w.Rollback
    On Error Resume Next
    If Not s Is Nothing Then s.Close
    If Not reader Is Nothing Then reader.Close
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
