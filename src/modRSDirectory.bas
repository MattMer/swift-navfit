Attribute VB_Name = "modRSDirectory"
Option Compare Binary
Option Explicit

Public Function RSSeniorKeys() As Variant
    RSSeniorKeys = Array("ReportingSeniorName", "ReportingSeniorGrade", "ReportingSeniorDesignator", "ReportingSeniorTitle", "ReportingSeniorUIC", "ReportingSeniorSSN", "ReportingSeniorAddress1", "ReportingSeniorAddress2", "ReportingSeniorCity", "ReportingSeniorState", "ReportingSeniorZip", "ReportingSeniorPhone", "ReportingSeniorDSN", "ReportingSeniorLastName", "ReportingSeniorFirstInitial", "ReportingSeniorMiddleInitial")
End Function
'Details kept on the senior's record but not on reports: the e-mail is printed on
'summary letters only.
Public Function RSRecordOnlyKeys() As Variant
    RSRecordOnlyKeys = Array("ReportingSeniorEmail")
End Function
'Every key a senior record stores: the report fields, then the record-only ones.
Public Function RSRecordKeys() As Variant
    RSRecordKeys = Split(Join(RSSeniorKeys(), "|") & "|" & Join(RSRecordOnlyKeys(), "|"), "|")
End Function
'The report fields of a senior record, without the record-only details.
Public Function RSSeniorReportValues(ByVal seniorID As Long) As Object
    Dim key As Variant
    Set RSSeniorReportValues = RSSeniorValues(seniorID)
    For Each key In RSRecordOnlyKeys(): RSSeniorReportValues.Remove CStr(key): Next
End Function
Public Function RSSeniorValues(ByVal seniorID As Long) As Object
    Dim result As Object, key As Variant, r As DAO.Recordset
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In RSRecordKeys(): result.Add CStr(key), Null: Next
    If seniorID > 0 Then
        If DCount("*", "ReportingSeniorRecord", "SeniorID=" & seniorID) <> 1 Then Err.Raise vbObjectError + 4800, "modRSDirectory.RSSeniorValues", "Reporting senior record does not exist."
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
'Reporting senior keys a command member profile can supply. The billet title,
'address and phone stay on the senior record: the profile holds the member's own.
Public Function RSMemberKeys() As Variant
    RSMemberKeys = Array("ReportingSeniorName", "ReportingSeniorLastName", "ReportingSeniorFirstInitial", "ReportingSeniorMiddleInitial", "ReportingSeniorGrade", "ReportingSeniorDesignator", "ReportingSeniorUIC", "ReportingSeniorSSN")
End Function
'Values a linked member supplies. Missing profile values stay Null so an update
'never clears what was typed on the senior record.
Public Function RSMemberValues(ByVal memberID As Long) As Object
    Dim result As Object, r As DAO.Recordset, ruad As Object, key As Variant, last As String, rest As String, parts As Variant, part As Variant, middle As String, pos As Long
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In RSMemberKeys(): result.Add CStr(key), Null: Next
    Set r = CurrentDb.OpenRecordset("SELECT FullName,ReportName,RankRate,MemberSSN,RUADData FROM CommandMember WHERE MemberID=" & memberID, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 4815, "modRSDirectory.RSMemberValues", "Command member does not exist."
    rest = Trim$(Nz(r!ReportName, "")): pos = InStr(rest, ",")
    If pos > 0 Then
        last = Trim$(Left$(rest, pos - 1)): rest = Trim$(Mid$(rest, pos + 1))
    Else
        rest = Trim$(Nz(r!FullName, "")): pos = InStrRev(rest, " ")
        If pos > 0 Then last = Mid$(rest, pos + 1): rest = Trim$(Left$(rest, pos - 1)) Else last = rest: rest = ""
    End If
    If Len(last) > 0 Then
        result("ReportingSeniorLastName") = last
        parts = Split(rest, " ")
        For Each part In parts
            If Len(part) > 0 Then
                If IsNull(result("ReportingSeniorFirstInitial")) Then result("ReportingSeniorFirstInitial") = UCase$(Left$(part, 1)) Else middle = middle & UCase$(Left$(part, 1))
            End If
        Next
        If Len(middle) > 0 Then result("ReportingSeniorMiddleInitial") = Left$(middle, 3)
        result("ReportingSeniorName") = last & IIf(IsNull(result("ReportingSeniorFirstInitial")), "", ", " & result("ReportingSeniorFirstInitial") & IIf(Len(middle) > 0, " " & Left$(middle, 3), ""))
    End If
    If Len(Trim$(Nz(r!RankRate, ""))) > 0 Then result("ReportingSeniorGrade") = Trim$(r!RankRate)
    If Len(Trim$(Nz(r!MemberSSN, ""))) > 0 Then result("ReportingSeniorSSN") = Trim$(r!MemberSSN)
    Set ruad = RUADDecode(r!RUADData.Value): r.Close
    If Len(Trim$(Nz(ruad("DESG"), ""))) > 0 Then result("ReportingSeniorDesignator") = Trim$(ruad("DESG"))
    If Len(Trim$(Nz(ruad("UIC"), ""))) > 0 Then result("ReportingSeniorUIC") = Trim$(ruad("UIC"))
    Set RSMemberValues = result
End Function
'Labels of profile values that differ from the given senior values.
Public Function RSMemberDifferences(ByVal memberID As Long, ByVal values As Object) As String
    Dim member As Object, key As Variant, text As String
    Set member = RSMemberValues(memberID)
    For Each key In RSMemberKeys()
        If Not IsNull(member(CStr(key))) And key <> "ReportingSeniorLastName" And key <> "ReportingSeniorFirstInitial" And key <> "ReportingSeniorMiddleInitial" Then
            If Not SameFieldValue(member(CStr(key)), values(CStr(key))) Then text = text & IIf(Len(text) > 0, ", ", "") & UXFieldLabel(CStr(key))
        End If
    Next
    RSMemberDifferences = text
End Function
'A member backs at most one senior record; a changed link needs an active member.
Public Sub RSValidateMemberLink(ByVal seniorID As Long, ByVal memberID As Variant)
    If IsNull(memberID) Then Exit Sub
    If seniorID > 0 Then
        If Nz(MajorScalar("SELECT MemberID FROM ReportingSeniorRecord WHERE SeniorID=" & seniorID), 0) = CLng(memberID) Then Exit Sub
    End If
    If Nz(MajorScalar("SELECT IsArchived FROM CommandMember WHERE MemberID=" & CLng(memberID)), True) Then Err.Raise vbObjectError + 4816, "modRSDirectory.RSValidateMemberLink", "Choose an active command member."
    If DCount("*", "ReportingSeniorRecord", "MemberID=" & CLng(memberID) & " AND SeniorID<>" & seniorID) > 0 Then Err.Raise vbObjectError + 4817, "modRSDirectory.RSValidateMemberLink", "This member is already linked to another reporting senior record."
End Sub
Public Sub RSValidateSenior(ByVal title As String, ByVal values As Object)
    Dim key As Variant, value As Variant, size As Long, i As Long
    If Len(Trim$(title)) = 0 Or Len(title) > 120 Then Err.Raise vbObjectError + 4801, "modRSDirectory.RSValidateSenior", "Enter a record name of up to 120 characters."
    'Callers that pass report fields only leave the e-mail unentered; the values they
    'pass are not changed, because some reuse them as report fields.
    value = Null
    If values.Exists("ReportingSeniorEmail") Then value = values("ReportingSeniorEmail")
    If Not IsNull(value) Then
        If VarType(value) <> vbString Then Err.Raise vbObjectError + 4803, "modRSDirectory.RSValidateSenior", "Reporting senior values must be text."
        value = Trim$(value)
        If Len(value) > 100 Then Err.Raise vbObjectError + 4804, "modRSDirectory.RSValidateSenior", "E-mail exceeds 100 characters."
        If Len(value) > 0 And (InStr(value, "@") < 2 Or InStr(value, " ") > 0 Or InStrRev(value, ".") < InStr(value, "@")) Then Err.Raise vbObjectError + 4804, "modRSDirectory.RSValidateSenior", "Enter the e-mail as name@domain, or leave it blank."
        If Len(value) = 0 Then values("ReportingSeniorEmail") = Null Else values("ReportingSeniorEmail") = value
    End If
    For Each key In RSSeniorKeys()
        If Not values.Exists(CStr(key)) Then Err.Raise vbObjectError + 4802, "modRSDirectory.RSValidateSenior", "Missing reporting senior value: " & key
        value = values(CStr(key))
        If Not IsNull(value) Then
            If VarType(value) <> vbString Then Err.Raise vbObjectError + 4803, "modRSDirectory.RSValidateSenior", "Reporting senior values must be text."
            size = CurrentDb.TableDefs("Report").Fields(CStr(key)).Size
            If Len(value) > size Then Err.Raise vbObjectError + 4804, "modRSDirectory.RSValidateSenior", UXFieldLabel(CStr(key)) & " exceeds " & size & " characters."
            If key = "ReportingSeniorSSN" And Len(value) > 0 Then
                If Len(value) <> 9 Then Err.Raise vbObjectError + 4805, "modRSDirectory.RSValidateSenior", "Reporting senior SSN requires nine digits."
                For i = 1 To 9
                    If Mid$(value, i, 1) < "0" Or Mid$(value, i, 1) > "9" Then Err.Raise vbObjectError + 4805, "modRSDirectory.RSValidateSenior", "Reporting senior SSN requires nine digits."
                Next
            End If
        End If
    Next
End Sub
'memberID: omit to keep the current directory link, Null to unlink.
Public Function RSSaveSenior(ByVal seniorID As Long, ByVal title As String, ByVal values As Object, Optional ByVal expectedVersion As Long = 0, Optional ByVal memberID As Variant) As Long
    Dim w As DAO.Workspace, db As DAO.Database, r As DAO.Recordset, key As Variant, author As Long, operation As Long, failure As Long, message As String, link As String, value As Variant
    RSValidateSenior title, values: author = MajorAuthor()
    If Not IsMissing(memberID) Then
        RSValidateMemberLink seniorID, memberID
        link = ",MemberID=" & IIf(IsNull(memberID), "Null", CStr(Nz(memberID, 0)))
    End If
    Set w = DBEngine.Workspaces(0): Set db = CurrentDb: w.BeginTrans
    On Error GoTo Failed
    If seniorID = 0 Then
        Set r = db.OpenRecordset("ReportingSeniorRecord", dbOpenDynaset)
        r.AddNew: r!SeniorGuid = NewIdentity(): r!Title = title: r!CreatedBy = author: r!ModifiedAt = Now(): r!VersionCounter = 1
        If Not IsMissing(memberID) Then r!MemberID = memberID
        r.Update: r.Bookmark = r.LastModified: seniorID = r!SeniorID: r.Close
    Else
        If expectedVersion < 1 Then Err.Raise vbObjectError + 4806, "modRSDirectory.RSSaveSenior", "Reload the reporting senior before saving."
        db.Execute "UPDATE ReportingSeniorRecord SET Title=" & QuoteSQL(title) & link & ",ModifiedAt=Now(),VersionCounter=VersionCounter+1 WHERE SeniorID=" & seniorID & " AND VersionCounter=" & expectedVersion, dbFailOnError
        If db.RecordsAffected <> 1 Then Err.Raise vbObjectError + 4806, "modRSDirectory.RSSaveSenior", "Reporting senior changed elsewhere. Reload before saving; pending values are retained."
        db.Execute "DELETE FROM ReportingSeniorValue WHERE SeniorID=" & seniorID, dbFailOnError
    End If
    Set r = db.OpenRecordset("ReportingSeniorValue", dbOpenDynaset)
    For Each key In RSRecordKeys()
        value = Null
        If values.Exists(CStr(key)) Then value = values(CStr(key))
        r.AddNew: r!SeniorID = seniorID: r!FieldKey = key: r!ValueText = value: r!IsNullValue = IsNull(value): r.Update
    Next
    r.Close
    operation = BeginOperation("Save reporting senior", "Reporting senior record " & seniorID)
    FinishOperation operation, "Completed", "Directory details saved; report values unchanged."
    w.CommitTrans: UXDirectorySourceInvalidate: RSSaveSenior = seniorID: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRSDirectory", "RSSaveSenior", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary: If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function
Public Sub RSValidateBaseline(ByVal paygrade As String, ByVal average As Variant, ByVal traitCount As Long, ByVal traitSum As Variant)
    Dim allowed As Boolean, grade As Variant
    For Each grade In RSTableGrades()
        If grade = paygrade Then allowed = True
    Next
    If Not allowed Then Err.Raise vbObjectError + 4807, "modRSDirectory.RSValidateBaseline", "Choose an evaluated paygrade."
    If traitCount < 0 Then Err.Raise vbObjectError + 4808, "modRSDirectory.RSValidateBaseline", "Previous trait count cannot be negative."
    If traitCount > 0 Or Not IsNull(average) Then
        If IsNull(average) Then Err.Raise vbObjectError + 4809, "modRSDirectory.RSValidateBaseline", "Enter the historical RSCA."
        If Not IsNumeric(average) Then Err.Raise vbObjectError + 4809, "modRSDirectory.RSValidateBaseline", "Enter a numeric RSCA."
        If CDbl(average) < 1 Or CDbl(average) > 5 Then Err.Raise vbObjectError + 4809, "modRSDirectory.RSValidateBaseline", "RSCA must be from 1.00 to 5.00."
    End If
    If Not IsNull(traitSum) Then
        If Not IsNumeric(traitSum) Then Err.Raise vbObjectError + 4810, "modRSDirectory.RSValidateBaseline", "Enter an exact numeric trait sum or leave it blank."
        If CDbl(traitSum) < traitCount Or CDbl(traitSum) > CDbl(traitCount) * 5 Or CDbl(traitSum) <> Fix(CDbl(traitSum)) Then Err.Raise vbObjectError + 4810, "modRSDirectory.RSValidateBaseline", "Exact trait sum must be a whole number between trait count and five times trait count."
        If traitCount > 0 Then
            If Int(CDbl(traitSum) / traitCount * 100 + 0.5) <> Int(CDbl(average) * 100 + 0.5) Then Err.Raise vbObjectError + 4811, "modRSDirectory.RSValidateBaseline", "Exact sum and count must agree with RSCA rounded to two decimals."
        End If
    End If
End Sub
Public Function RSAddBaseline(ByVal seniorID As Long, ByVal paygrade As String, ByVal asOfDate As Date, ByVal average As Variant, ByVal traitCount As Long, ByVal traitSum As Variant, Optional ByVal note As String = "") As Long
    Dim r As DAO.Recordset, w As DAO.Workspace, operation As Long, failure As Long, message As String
    RSValidateBaseline paygrade, average, traitCount, traitSum
    If Len(note) > 255 Then Err.Raise vbObjectError + 4812, "modRSDirectory.RSAddBaseline", "Source note must be 255 characters or fewer."
    If DCount("*", "ReportingSeniorRecord", "SeniorID=" & seniorID) <> 1 Then Err.Raise vbObjectError + 4800, "modRSDirectory.RSAddBaseline", "Choose a saved reporting senior record."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("RSCABaseline", dbOpenDynaset)
    r.AddNew: r!SeniorID = seniorID: r!Paygrade = paygrade: r!AsOfDate = DateValue(asOfDate): r!CurrentRSCA = average: r!TraitCount = traitCount: r!TraitSum = traitSum: r!SourceNote = note: r!CreatedBy = MajorAuthor(): r!CreatedAt = Now()
    r.Update: r.Bookmark = r.LastModified: RSAddBaseline = r!BaselineID: r.Close
    operation = BeginOperation("Record historical RSCA", "Reporting senior " & seniorID & ", paygrade " & paygrade & ", baseline " & RSAddBaseline)
    FinishOperation operation, "Completed", "Historical totals retained; group contributions are not added to this baseline automatically."
    w.CommitTrans: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRSDirectory", "RSAddBaseline", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary: If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function
Public Sub RSSetGroupBaseline(ByVal groupID As Long, ByVal seniorID As Long, ByVal baselineID As Long, ByVal excludesGroup As Boolean)
    Dim r As DAO.Recordset, w As DAO.Workspace, operation As Long, failure As Long, message As String
    RSGuardPostedSenior groupID, seniorID
    If Nz(MajorScalar("SELECT IsArchived FROM SummaryGroup WHERE GroupID=" & groupID), True) Then Err.Raise vbObjectError + 4813, "modRSDirectory.RSSetGroupBaseline", "Choose an active summary group."
    If DCount("*", "ReportingSeniorRecord", "SeniorID=" & seniorID) <> 1 Then Err.Raise vbObjectError + 4800, "modRSDirectory.RSSetGroupBaseline", "Choose a reporting senior."
    If baselineID > 0 Then
        If DCount("*", "RSCABaseline", "BaselineID=" & baselineID & " AND SeniorID=" & seniorID & " AND Paygrade=" & QuoteSQL(CStr(MajorScalar("SELECT Paygrade FROM SummaryGroup WHERE GroupID=" & groupID)))) <> 1 Then Err.Raise vbObjectError + 4814, "modRSDirectory.RSSetGroupBaseline", "Baseline senior and evaluated paygrade must match this group."
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
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRSDirectory", "RSSetGroupBaseline", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary: If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Sub
Public Function RSGroupRows(ByVal groupID As Long) As DAO.Recordset
    Dim db As DAO.Database, rows As DAO.Recordset
    Set db = CurrentDb
    Set rows = db.OpenRecordset("SELECT r.* FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & groupID, dbOpenSnapshot)
    AssertReportCacheRows rows
    If Not rows.BOF Or Not rows.EOF Then rows.MoveFirst
    Set RSGroupRows = rows
End Function
Public Function RSGroupPSG(ByVal groupID As Long, Optional ByVal loaded As DAO.Recordset = Nothing) As Object
    Dim baselineID As Long, result As Object, r As DAO.Recordset, members As DAO.Recordset, senior As Object, values As Object, key As Variant, count As Long, total As Double, previousCount As Long, previousSum As Double, memberCount As Long
    On Error GoTo AccountingUnavailable
    Set result = CreateObject("Scripting.Dictionary"): result.Add "Value", Null: result.Add "Reason", "Choose a reporting senior and historical baseline.": result.Add "Estimated", False: result.Add "Partial", False
    baselineID = Nz(DLookup("BaselineID", "GroupRSCABaseline", "GroupID=" & groupID), 0)
    baselineID = RSCalculationBaseline(groupID, baselineID)
    Set r = CurrentDb.OpenRecordset("SELECT b.*,g.ExcludesGroup,g.SeniorID AS SelectedSeniorID FROM GroupRSCABaseline AS g, RSCABaseline AS b WHERE b.BaselineID=" & baselineID & " AND g.GroupID=" & groupID, dbOpenSnapshot)
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
    result.Add "SeniorID", CLng(r!SeniorID): result.Add "BaselineID", CLng(r!BaselineID)
    r.Close
    RSExcludePostedGroup groupID, baselineID, previousSum, previousCount
    result.Add "HistoricalSum", previousSum: result.Add "HistoricalCount", previousCount
    If loaded Is Nothing Then Set members = RSGroupRows(groupID) Else Set members = loaded.Clone
    If Not (members.BOF And members.EOF) Then members.MoveFirst
    Do While Not members.EOF
        If Len(Nz(senior("ReportingSeniorSSN"), "")) = 0 Or Nz(members!ReportingSeniorSSN, "") <> Nz(senior("ReportingSeniorSSN"), "") Then
            result("Reason") = "All group reports must use the selected reporting senior SSN. Apply senior details first.": members.Close: Set RSGroupPSG = result: Exit Function
        End If
        If Nz(members!IsArchived, True) Then
            result("Reason") = "Remove or recover archived group members before calculating PSG RSCA.": members.Close: Set RSGroupPSG = result: Exit Function
        End If
        If Nz(members!ObservationStatus, "") <> "NOB" Then
            memberCount = 0
            For Each key In TraitKeys()
                Select Case Nz(members.Fields(CStr(key)).Value, "")
                    Case "1.0", "2.0", "3.0", "4.0", "5.0": total = total + CDbl(members.Fields(CStr(key)).Value): count = count + 1: memberCount = memberCount + 1
                End Select
            Next
            If memberCount < 7 Then result("Partial") = True
        End If
        members.MoveNext
    Loop
    members.Close
    result.Add "GroupSum", total: result.Add "GroupCount", count
    If count = 0 Then result("Reason") = "No graded traits are available.": Set RSGroupPSG = result: Exit Function
    result("Value") = (previousSum + total) / (CDbl(previousCount) + count)
    result("Reason") = "(" & Format$(previousSum, "0.00") & " historical sum + " & total & " group sum) / (" & previousCount & " historical traits + " & count & " group traits)"
    Set RSGroupPSG = result
    Exit Function
AccountingUnavailable:
    If Err.Number <> vbObjectError + 4851 And Err.Number <> vbObjectError + 4852 Then Err.Raise Err.Number, "RSGroupPSG", Err.Description
    result("Value") = Null: result("Reason") = Err.Description
    Set RSGroupPSG = result
End Function
