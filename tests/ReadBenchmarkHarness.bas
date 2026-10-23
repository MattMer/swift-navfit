Attribute VB_Name = "ReadBenchmarkHarness"
Option Explicit
Public Function SeedReadBenchmark(ByVal size As Long) As String
    Dim i As Long, id As Long, group As Long, senior As Long, baseline As Long, member As Long, folders(1 To 3) As Long
    Dim values As Object, traits As Object, key As Variant, marks As Variant, observation As String, ignored As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC BATCHED READS", "Read tester"
    folders(1) = CreateFolder("SYNTHETIC root")
    folders(2) = CreateFolder("SYNTHETIC child", folders(1))
    folders(3) = CreateFolder("SYNTHETIC leaf", folders(2))
    group = CreateSummaryGroup("SYNTHETIC large group", "EVAL", "E5")
    Set values = RSSeniorValues(0): values("ReportingSeniorSSN") = "000999999": values("ReportingSeniorName") = "SYNTHETIC RS"
    senior = RSSaveSenior(0, "SYNTHETIC senior", values)
    baseline = RSAddBaseline(senior, "E5", #1/1/2026#, 4, 100, 400, "SYNTHETIC exact historical totals")
    RSSetGroupBaseline group, senior, baseline, True
    For i = 1 To size
        marks = Array("4.0", "4.0", "4.0", "4.0", "4.0", "4.0", "4.0"): observation = "Observed"
        If i Mod 7 = 0 Then observation = "NOB"
        If i Mod 7 = 1 Then marks(1) = Null
        If i Mod 7 = 2 Then marks(1) = "NOB"
        id = CommitDraft(0, "EVAL", "SYNTHETIC " & Format$(i, "0000"), "SYNTHETIC " & Format$(i, "0000"), #1/1/2026#, #9/30/2026#, "SYNTHETIC private notes", "Read benchmark seed", "SYNTHETIC official comments", administration:=Array("SYNTHETIC " & Format$(i, "0000"), "YN2", Null, Null, Null), traits:=marks, assessment:=Array(observation, "Promotable"), reportingSenior:=FieldsSubset(values, ReportingSeniorKeys()), memberDetails:=Array("000" & Format$(i, "000000"), Null, "Yes", "No", "No", "No", "2026-01-01"))
        If i Mod 3 <> 0 Then AssignFolder id, folders(3)
        AddGroupMember group, id, "E5"
    Next
    member = CLng(DLookup("MemberID", "MemberReport", "ReportID=1"))
    Set traits = CreateObject("Scripting.Dictionary")
    For Each key In TraitKeys(): traits.Add CStr(key), "4.0": Next
    For i = 1 To size
        observation = "Observed": traits("Trait33") = "4.0"
        If i Mod 7 = 0 Then observation = "NOB"
        If i Mod 7 = 1 Then traits("Trait33") = Null
        If i Mod 7 = 2 Then traits("Trait33") = "NOB"
        ignored = SavePriorEvaluation(0, 0, member, "EVAL", "YN2", DateSerial(2010, 1, i), DateSerial(2010, 1, i), "Promotable", observation, traits, "SYNTHETIC history")
    Next
    SeedReadBenchmark = "PASS: " & size & " reports and prior records": Exit Function
Failed:
    SeedReadBenchmark = "FAIL: " & Err.Number & ": " & Err.Description
End Function
Private Function Cell(ByVal value As Variant) As String
    If IsNull(value) Then Cell = "N;" Else Cell = Len(CStr(value)) & ":" & CStr(value) & ";"
End Function
Private Function ListText(ByVal list As Access.ListBox) As String
    Dim row As Long, col As Long
    For row = 0 To list.ListCount - 1
        For col = 0 To list.ColumnCount - 1: ListText = ListText & Cell(list.Column(col, row)): Next
        ListText = ListText & vbLf
    Next
End Function
Public Function PrepareReadViews() As Boolean
    DoCmd.OpenForm "frmBulkFolders"
    DoCmd.OpenForm "frmUXDirectory"
    Forms!frmUXDirectory!lstMembers = DLookup("MemberID", "MemberReport", "ReportID=1")
    PrepareReadViews = UXDirectoryHistory()
End Function
Public Function ReadBenchmarkCase(ByVal kind As String) As String
    Dim r As DAO.Recordset, selection As String, result As Object, key As Variant
    Select Case kind
    Case "Folder stamp"
        Set r = CurrentDb.OpenRecordset("SELECT ReportID FROM Report ORDER BY ReportID", dbOpenSnapshot)
        Do While Not r.EOF
            If selection <> "" Then selection = selection & ","
            selection = selection & r!ReportID: r.MoveNext
        Loop
        r.Close: ReadBenchmarkCase = FolderBatchStamp(selection)
    Case "Folder display"
        Call LoadBulkFolders
        ReadBenchmarkCase = ListText(Forms!frmBulkFolders!lstReports)
    Case "Directory history"
        Call UXHistoryLoad
        ReadBenchmarkCase = ListText(Forms!frmUXMemberHistory!lstHistory)
    Case "Group RSCA"
        Set r = RSGroupRows(1): Set result = RSGroupPSG(1, r): r.Close
        For Each key In Array("Value", "Reason", "Estimated", "Partial", "HistoricalSum", "HistoricalCount")
            ReadBenchmarkCase = ReadBenchmarkCase & CStr(key) & "=" & Cell(result(key))
        Next
    Case Else
        Err.Raise 5, , "Unknown benchmark case."
    End Select
End Function
Public Function ReadHistoryCounts() As String
    Dim table As Variant
    For Each table In Array("Report", "ReportRevision", "RevisionValue", "RevisionDate", "RevisionChoice", "FolderMove", "GroupSnapshot", "ComparisonSet", "OperationJournal", "RecoveryDraft", "PriorEvaluation", "PriorEvaluationTrait", "GroupCompletion", "RSCABaseline")
        ReadHistoryCounts = ReadHistoryCounts & table & ":" & DCount("*", CStr(table)) & ";"
    Next
End Function
