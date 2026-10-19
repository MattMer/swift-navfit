Attribute VB_Name = "GroupReadScaleHarness"
Option Explicit
Private Function SampleGroup(ByVal g As Long, ByVal legacy As Boolean, ByRef resultText As String) As Double
    Dim started As Double, average As Variant, result As Object, rows As DAO.Recordset, details As String, key As Variant
    started = Timer
    If legacy Then
        average = LegacyUXGroupAverage(g): Set result = LegacyRSGroupPSG(g): details = LegacyUXGroupCalculationText(g)
    Else
        average = UXGroupAverage(g): Set rows = RSGroupRows(g)
        Set result = RSGroupPSG(g, rows): details = UXGroupCalculationText(g, rows): rows.Close
    End If
    resultText = Nz(average, "NULL") & "|" & details
    For Each key In result.Keys: resultText = resultText & "|" & key & "=" & Nz(result(key), "NULL"): Next
    SampleGroup = Timer - started
    If SampleGroup < 0 Then SampleGroup = SampleGroup + 86400
    SampleGroup = SampleGroup * 1000
End Function
Public Function TestGroupReadScale() As String
    Dim target As Variant, i As Long, g As Long, id As Long, senior As Long, baseline As Long, values As Object, trial As Long, oldText As String, freshText As String, oldTime As Double, freshTime As Double, output As String, revisions As Long, snapshots As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC GROUP READ SCALE", "Tester"
    g = CreateSummaryGroup("Scale group", "EVAL", "E5")
    Set values = RSSeniorValues(0): values("ReportingSeniorSSN") = "000123456": values("ReportingSeniorName") = "SYNTHETIC RS"
    senior = RSSaveSenior(0, "Scale senior", values)
    baseline = RSAddBaseline(senior, "E5", #1/1/2026#, 4, 100, 400, "Exact synthetic baseline")
    RSSetGroupBaseline g, senior, baseline, True
    output = "Members,Trial,LegacyMs,PreparedMs" & vbCrLf
    For Each target In Array(25, 50, 100)
        For i = DCount("*", "Report") + 1 To CLng(target)
            id = CommitDraft(0, "EVAL", "Scale report", "Member " & i, Null, Null, "")
            CurrentDb.Execute "UPDATE Report SET Trait33='4.0',Trait34='4.0',Trait35='4.0',Trait36='4.0',Trait37='4.0',Trait38='4.0',Trait39='4.0',ReportingSeniorSSN='000123456' WHERE ReportID=" & id, dbFailOnError
            CurrentDb.Execute "UPDATE RevisionValue SET ValueText='4.0',IsNullValue=False WHERE RevisionID=" & CurrentRevision(id) & " AND FieldKey In ('Trait33','Trait34','Trait35','Trait36','Trait37','Trait38','Trait39')", dbFailOnError
            CurrentDb.Execute "UPDATE RevisionValue SET ValueText='000123456',IsNullValue=False WHERE RevisionID=" & CurrentRevision(id) & " AND FieldKey='ReportingSeniorSSN'", dbFailOnError
            AddGroupMember g, id, "E5"
        Next
        revisions = DCount("*", "ReportRevision"): snapshots = DCount("*", "GroupSnapshot") + DCount("*", "ComparisonSet")
        For trial = 1 To 5
            If trial Mod 2 = 1 Then
                oldTime = SampleGroup(g, True, oldText): freshTime = SampleGroup(g, False, freshText)
            Else
                freshTime = SampleGroup(g, False, freshText): oldTime = SampleGroup(g, True, oldText)
            End If
            If StrComp(oldText, freshText, vbBinaryCompare) <> 0 Then Err.Raise 5, , "Calculation result or explanation differs"
            output = output & target & "," & trial & "," & Format$(oldTime, "0.0") & "," & Format$(freshTime, "0.0") & vbCrLf
        Next
        If DCount("*", "ReportRevision") <> revisions Or DCount("*", "GroupSnapshot") + DCount("*", "ComparisonSet") <> snapshots Then Err.Raise 5, , "Calculation created history"
    Next
    TestGroupReadScale = "PASS:" & vbCrLf & output: Exit Function
Failed:
    TestGroupReadScale = "FAIL: " & Err.Number & " " & Err.Description
End Function
