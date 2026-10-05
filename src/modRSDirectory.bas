Attribute VB_Name = "modRSDirectory"
Option Compare Binary
Option Explicit

Public Function RSSeniorKeys() As Variant
    RSSeniorKeys = Array("ReportingSeniorName", "ReportingSeniorGrade", "ReportingSeniorDesignator", "ReportingSeniorTitle", "ReportingSeniorUIC", "ReportingSeniorSSN", "ReportingSeniorAddress1", "ReportingSeniorAddress2", "ReportingSeniorCity", "ReportingSeniorState", "ReportingSeniorZip", "ReportingSeniorPhone", "ReportingSeniorDSN", "ReportingSeniorLastName", "ReportingSeniorFirstInitial", "ReportingSeniorMiddleInitial")
End Function
Public Function RSSeniorValues(ByVal seniorID As Long) As Object
    Dim result As Object, key As Variant, r As DAO.Recordset
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In RSSeniorKeys(): result.Add CStr(key), Null: Next
    If seniorID > 0 Then
        If DCount("*", "ReportingSeniorRecord", "SeniorID=" & seniorID) <> 1 Then Err.Raise vbObjectError + 4800, , "Reporting senior record does not exist."
        Set r = CurrentDb.OpenRecordset("SELECT * FROM ReportingSeniorValue WHERE SeniorID=" & seniorID, dbOpenSnapshot)
        Do While Not r.EOF
            If result.Exists(CStr(r!FieldKey)) Then
                If Not r!IsNullValue Then result(CStr(r!FieldKey)) = r!ValueText.Value
            End If
            r.MoveNext
        Loop
        r.Close
    End If
    Set RSSeniorValues = result
End Function
Public Sub RSValidateSenior(ByVal title As String, ByVal values As Object)
    Dim key As Variant, value As Variant, size As Long, i As Long
    If Len(Trim$(title)) = 0 Or Len(title) > 120 Then Err.Raise vbObjectError + 4801, , "Enter a record name of up to 120 characters."
    For Each key In RSSeniorKeys()
        If Not values.Exists(CStr(key)) Then Err.Raise vbObjectError + 4802, , "Missing reporting senior value: " & key
        value = values(CStr(key))
        If Not IsNull(value) Then
            If VarType(value) <> vbString Then Err.Raise vbObjectError + 4803, , "Reporting senior values must be text."
            size = CurrentDb.TableDefs("Report").Fields(CStr(key)).Size
            If Len(value) > size Then Err.Raise vbObjectError + 4804, , UXFieldLabel(CStr(key)) & " exceeds " & size & " characters."
            If key = "ReportingSeniorSSN" And Len(value) > 0 Then
                If Len(value) <> 9 Then Err.Raise vbObjectError + 4805, , "Reporting senior SSN requires nine digits."
                For i = 1 To 9
                    If Mid$(value, i, 1) < "0" Or Mid$(value, i, 1) > "9" Then Err.Raise vbObjectError + 4805, , "Reporting senior SSN requires nine digits."
                Next
            End If
        End If
    Next
End Sub
Public Function RSSaveSenior(ByVal seniorID As Long, ByVal title As String, ByVal values As Object, Optional ByVal expectedVersion As Long = 0) As Long
    Dim w As DAO.Workspace, db As DAO.Database, r As DAO.Recordset, key As Variant, author As Long, operation As Long, failure As Long, message As String
    RSValidateSenior title, values: author = MajorAuthor()
    Set w = DBEngine.Workspaces(0): Set db = CurrentDb: w.BeginTrans
    On Error GoTo Failed
    If seniorID = 0 Then
        Set r = db.OpenRecordset("ReportingSeniorRecord", dbOpenDynaset)
        r.AddNew: r!SeniorGuid = NewIdentity(): r!Title = title: r!CreatedBy = author: r!ModifiedAt = Now(): r!VersionCounter = 1
        r.Update: r.Bookmark = r.LastModified: seniorID = r!SeniorID: r.Close
    Else
        If expectedVersion < 1 Then Err.Raise vbObjectError + 4806, , "Reload the reporting senior before saving."
        db.Execute "UPDATE ReportingSeniorRecord SET Title=" & QuoteSQL(title) & ",ModifiedAt=Now(),VersionCounter=VersionCounter+1 WHERE SeniorID=" & seniorID & " AND VersionCounter=" & expectedVersion, dbFailOnError
        If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 4806, , "Reporting senior changed elsewhere. Reload before saving; pending values are retained."
        db.Execute "DELETE FROM ReportingSeniorValue WHERE SeniorID=" & seniorID, dbFailOnError
    End If
    Set r = db.OpenRecordset("ReportingSeniorValue", dbOpenDynaset)
    For Each key In RSSeniorKeys()
        r.AddNew: r!SeniorID = seniorID: r!FieldKey = key: r!ValueText = values(CStr(key)): r!IsNullValue = IsNull(values(CStr(key))): r.Update
    Next
    r.Close
    operation = BeginOperation("Save reporting senior", "Reporting senior record " & seniorID)
    FinishOperation operation, "Completed", "Directory details saved; report values unchanged."
    w.CommitTrans: RSSaveSenior = seniorID: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next: w.Rollback: If Not r Is Nothing Then r.Close
    On Error GoTo 0: Err.Raise failure, "RSSaveSenior", message
End Function
Public Sub RSValidateBaseline(ByVal paygrade As String, ByVal average As Variant, ByVal traitCount As Long, ByVal traitSum As Variant)
    Dim allowed As Boolean, grade As Variant
    For Each grade In Array("E1", "E2", "E3", "E4", "E5", "E6", "E7", "E8", "E9", "W1", "W2", "W3", "W4", "W5", "O1", "O2", "O3", "O4", "O5", "O6")
        If grade = paygrade Then allowed = True
    Next
    If Not allowed Then Err.Raise vbObjectError + 4807, , "Choose an evaluated paygrade."
    If traitCount < 0 Then Err.Raise vbObjectError + 4808, , "Previous trait count cannot be negative."
    If traitCount > 0 Or Not IsNull(average) Then
        If IsNull(average) Then Err.Raise vbObjectError + 4809, , "Enter the historical RSCA."
        If Not IsNumeric(average) Then Err.Raise vbObjectError + 4809, , "Enter a numeric RSCA."
        If CDbl(average) < 1 Or CDbl(average) > 5 Then Err.Raise vbObjectError + 4809, , "RSCA must be from 1.00 to 5.00."
    End If
    If Not IsNull(traitSum) Then
        If Not IsNumeric(traitSum) Then Err.Raise vbObjectError + 4810, , "Enter an exact numeric trait sum or leave it blank."
        If CDbl(traitSum) < traitCount Or CDbl(traitSum) > CDbl(traitCount) * 5 Or CDbl(traitSum) <> Fix(CDbl(traitSum)) Then Err.Raise vbObjectError + 4810, , "Exact trait sum must be a whole number between trait count and five times trait count."
        If traitCount > 0 Then
            If Int(CDbl(traitSum) / traitCount * 100 + 0.5) <> Int(CDbl(average) * 100 + 0.5) Then Err.Raise vbObjectError + 4811, , "Exact sum and count must agree with RSCA rounded to two decimals."
        End If
    End If
End Sub
Public Function RSAddBaseline(ByVal seniorID As Long, ByVal paygrade As String, ByVal asOfDate As Date, ByVal average As Variant, ByVal traitCount As Long, ByVal traitSum As Variant, Optional ByVal note As String = "") As Long
    Dim r As DAO.Recordset, w As DAO.Workspace, operation As Long, failure As Long, message As String
    RSValidateBaseline paygrade, average, traitCount, traitSum
    If Len(note) > 255 Then Err.Raise vbObjectError + 4812, , "Source note must be 255 characters or fewer."
    If DCount("*", "ReportingSeniorRecord", "SeniorID=" & seniorID) <> 1 Then Err.Raise vbObjectError + 4800, , "Choose a saved reporting senior record."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("RSCABaseline", dbOpenDynaset)
    r.AddNew: r!SeniorID = seniorID: r!Paygrade = paygrade: r!AsOfDate = DateValue(asOfDate): r!CurrentRSCA = average: r!TraitCount = traitCount: r!TraitSum = traitSum: r!SourceNote = note: r!CreatedBy = MajorAuthor(): r!CreatedAt = Now()
    r.Update: r.Bookmark = r.LastModified: RSAddBaseline = r!BaselineID: r.Close
    operation = BeginOperation("Record historical RSCA", "Reporting senior " & seniorID & ", paygrade " & paygrade & ", baseline " & RSAddBaseline)
    FinishOperation operation, "Completed", "Historical totals retained; group contributions are not added to this baseline automatically."
    w.CommitTrans: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next: w.Rollback: If Not r Is Nothing Then r.Close
    On Error GoTo 0: Err.Raise failure, "RSAddBaseline", message
End Function
Public Sub RSSetGroupBaseline(ByVal groupID As Long, ByVal seniorID As Long, ByVal baselineID As Long, ByVal excludesGroup As Boolean)
    Dim r As DAO.Recordset, w As DAO.Workspace, operation As Long, failure As Long, message As String
    If Nz(MajorScalar("SELECT IsArchived FROM SummaryGroup WHERE GroupID=" & groupID), True) Then Err.Raise vbObjectError + 4813, , "Choose an active summary group."
    If DCount("*", "ReportingSeniorRecord", "SeniorID=" & seniorID) <> 1 Then Err.Raise vbObjectError + 4800, , "Choose a reporting senior."
    If baselineID > 0 Then
        If DCount("*", "RSCABaseline", "BaselineID=" & baselineID & " AND SeniorID=" & seniorID & " AND Paygrade=" & QuoteSQL(CStr(MajorScalar("SELECT Paygrade FROM SummaryGroup WHERE GroupID=" & groupID)))) <> 1 Then Err.Raise vbObjectError + 4814, , "Baseline senior and evaluated paygrade must match this group."
    End If
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM GroupRSCABaseline WHERE GroupID=" & groupID, dbOpenDynaset)
    If r.EOF Then r.AddNew: r!GroupID = groupID Else r.Edit
    r!SeniorID = seniorID: If baselineID > 0 Then r!BaselineID = baselineID Else r!BaselineID = Null
    r!ExcludesGroup = excludesGroup: r.Update: r.Close
    operation = BeginOperation("Select group RSCA baseline", "Group " & groupID & "; senior " & seniorID & "; baseline " & baselineID)
    FinishOperation operation, "Completed", "Historical totals exclude current group: " & CStr(excludesGroup)
    w.CommitTrans: Exit Sub
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next: w.Rollback: If Not r Is Nothing Then r.Close
    On Error GoTo 0: Err.Raise failure, "RSSetGroupBaseline", message
End Sub
Public Function RSGroupPSG(ByVal groupID As Long) As Object
    Dim result As Object, r As DAO.Recordset, members As DAO.Recordset, senior As Object, values As Object, key As Variant, count As Long, total As Double, previousCount As Long, previousSum As Double, memberCount As Long
    Set result = CreateObject("Scripting.Dictionary"): result.Add "Value", Null: result.Add "Reason", "Choose a reporting senior and historical baseline.": result.Add "Estimated", False: result.Add "Partial", False
    Set r = CurrentDb.OpenRecordset("SELECT b.*,g.ExcludesGroup,g.SeniorID AS SelectedSeniorID FROM GroupRSCABaseline AS g INNER JOIN RSCABaseline AS b ON g.BaselineID=b.BaselineID WHERE g.GroupID=" & groupID, dbOpenSnapshot)
    If r.EOF Then r.Close: Set RSGroupPSG = result: Exit Function
    If Not r!ExcludesGroup Then result("Reason") = "Confirm that historical totals exclude the current group.": r.Close: Set RSGroupPSG = result: Exit Function
    If r!Paygrade <> MajorScalar("SELECT Paygrade FROM SummaryGroup WHERE GroupID=" & groupID) Then result("Reason") = "Historical baseline paygrade does not match this group.": r.Close: Set RSGroupPSG = result: Exit Function
    If r!SelectedSeniorID <> r!SeniorID Then result("Reason") = "Baseline belongs to a different reporting senior.": r.Close: Set RSGroupPSG = result: Exit Function
    Set senior = RSSeniorValues(r!SeniorID)
    previousCount = r!TraitCount
    RSValidateBaseline CStr(r!Paygrade), r!CurrentRSCA.Value, previousCount, r!TraitSum.Value
    If previousCount = 0 Then
        previousSum = 0
    ElseIf IsNull(r!TraitSum) Then
        previousSum = CDbl(r!CurrentRSCA) * previousCount: result("Estimated") = True
    Else
        previousSum = r!TraitSum
    End If
    r.Close
    Set members = CurrentDb.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & groupID, dbOpenSnapshot)
    Do While Not members.EOF
        Set values = SavedFields(members!ReportID)
        If Len(Nz(senior("ReportingSeniorSSN"), "")) = 0 Or Nz(values("ReportingSeniorSSN"), "") <> Nz(senior("ReportingSeniorSSN"), "") Then
            result("Reason") = "All group reports must use the selected reporting senior SSN. Apply senior details first.": members.Close: Set RSGroupPSG = result: Exit Function
        End If
        If Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & members!ReportID), True) Then
            result("Reason") = "Remove or recover archived group members before calculating PSG RSCA.": members.Close: Set RSGroupPSG = result: Exit Function
        End If
        If Nz(values("ObservationStatus"), "") <> "NOB" Then
            memberCount = 0
            For Each key In TraitKeys()
                Select Case Nz(values(CStr(key)), "")
                    Case "1.0", "2.0", "3.0", "4.0", "5.0": total = total + CDbl(values(CStr(key))): count = count + 1: memberCount = memberCount + 1
                End Select
            Next
            If memberCount < 7 Then result("Partial") = True
        End If
        members.MoveNext
    Loop
    members.Close
    If count = 0 Then result("Reason") = "No graded traits are available.": Set RSGroupPSG = result: Exit Function
    result("Value") = (previousSum + total) / (CDbl(previousCount) + count)
    result("Reason") = "(" & Format$(previousSum, "0.00") & " historical sum + " & total & " group sum) / (" & previousCount & " historical traits + " & count & " group traits)"
    Set RSGroupPSG = result
End Function
