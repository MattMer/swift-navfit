Attribute VB_Name = "modCounseling"
Option Compare Database
Option Explicit

Public Const COUNSELING_SOURCE As String = "BUPERSINST 1610.10H CH-1 (16 Dec 2025), enclosure (2) chapter 18, paragraph 18-6 midterm performance counseling schedules"
Public Const COUNSELING_COMPLETED As String = "Completed"
Public Const COUNSELING_NOT_PERFORMED As String = "Not performed"

'Schedule rows: effective from|effective to|worn paygrades|month. Active, TAR and
'Inactive share one month in both tables. The November 2026 captain row covers the
'shortened 31 Jul 2026 - 31 Mar 2027 period. Earlier history uses the first table.
Public Function CounselingScheduleRows() As Variant
    CounselingScheduleRows = Array( _
        "1900-01-01|2026-12-31|O6|1", "2026-11-01|2026-11-30|O6|11", "1900-01-01|2026-12-31|O5|10", "1900-01-01|2026-12-31|O4|4", _
        "1900-01-01|2026-12-31|O3|7", "1900-01-01|2026-12-31|O2|8", "1900-01-01|2026-12-31|O1|11", "1900-01-01|2026-12-31|W3,W4,W5|9", _
        "1900-01-01|2026-12-31|E9|10", "1900-01-01|2026-12-31|E7,E8|3", "1900-01-01|2026-12-31|E6|5", "1900-01-01|2026-12-31|E5|9", _
        "1900-01-01|2026-12-31|E4|12", "1900-01-01|2026-12-31|E1,E2,E3|1", _
        "2027-01-01|9999-12-31|O6|9", "2027-01-01|9999-12-31|O5|12", "2027-01-01|9999-12-31|O4|4", "2027-01-01|9999-12-31|O3|7", _
        "2027-01-01|9999-12-31|O2|8", "2027-01-01|9999-12-31|O1|11", "2027-01-01|9999-12-31|W3,W4,W5|9", "2027-01-01|9999-12-31|W1,W2|3", _
        "2027-01-01|9999-12-31|E9|10", "2027-01-01|9999-12-31|E7,E8|3", "2027-01-01|9999-12-31|E6|5", "2027-01-01|9999-12-31|E5|9", _
        "2027-01-01|9999-12-31|E4|12", "2027-01-01|9999-12-31|E1,E2,E3|1")
End Function

'Due months (as month-end dates) for one paygrade between two dates, ascending.
Public Function CounselingDueMonths(ByVal paygrade As String, ByVal fromDate As Date, ByVal toDate As Date) As Collection
    Dim result As New Collection, row As Variant, parts As Variant, y As Long, monthEnd As Date, i As Long, placed As Boolean
    For Each row In CounselingScheduleRows()
        parts = Split(CStr(row), "|")
        If InStr("," & parts(2) & ",", "," & paygrade & ",") > 0 Then
            For y = Year(fromDate) To Year(toDate)
                monthEnd = DateSerial(y, CLng(parts(3)) + 1, 0)
                If monthEnd >= fromDate And monthEnd <= toDate And monthEnd >= CDate(parts(0)) And DateSerial(y, CLng(parts(3)), 1) <= CDate(parts(1)) Then
                    placed = False
                    For i = 1 To result.Count
                        If result(i) = monthEnd Then placed = True: Exit For
                        If result(i) > monthEnd Then result.Add monthEnd, Before:=i: placed = True: Exit For
                    Next
                    If Not placed Then result.Add monthEnd
                End If
            Next
        End If
    Next
    Set CounselingDueMonths = result
End Function

'Status of one member's counseling on a date. Each due month owns the window from
'halfway after the previous due month to halfway before the next, so a session held
'earlier or later than scheduled still satisfies its cycle.
Public Function CounselingMemberStatus(ByVal memberID As Long, Optional ByVal asOf As Date = 0) As Object
    Dim result As Object, rank As String, grade As String, due As Collection, i As Long, lastIndex As Long, nextIndex As Long
    Dim windowStart As Date, windowEnd As Date, monthStart As Date, recorded As Variant
    If asOf = 0 Then asOf = Date
    Set result = CreateObject("Scripting.Dictionary")
    result("Status") = "": result("DueMonth") = Null: result("LastCounseled") = Null: result("Note") = ""
    result("LastCounseled") = MajorScalar("SELECT Max(CounseledOn) FROM MemberCounseling WHERE MemberID=" & memberID & " AND IsArchived=False AND Outcome='" & COUNSELING_COMPLETED & "'")
    rank = Nz(MajorScalar("SELECT RankRate FROM CommandMember WHERE MemberID=" & memberID), "")
    grade = ValidationWornPaygrade(rank)
    If grade = "" Then result("Status") = "Enter rank/rate": Set CounselingMemberStatus = result: Exit Function
    'Since NAVADMIN 132/26 (3 Jun 2026) E-9 and CWO5 have no Regular periodic report;
    'the command does not track their midterm counseling as due.
    If (grade = "E9" Or grade = "W5") And asOf >= #6/3/2026# Then result("Status") = "Not scheduled": result("Note") = "No Regular periodic report.": Set CounselingMemberStatus = result: Exit Function
    Set due = CounselingDueMonths(grade, DateAdd("yyyy", -3, asOf), DateAdd("yyyy", 2, asOf))
    If due.Count = 0 Then result("Status") = "Not scheduled": Set CounselingMemberStatus = result: Exit Function
    monthStart = DateSerial(Year(asOf), Month(asOf), 1)
    For i = 1 To due.Count
        If due(i) < monthStart Then lastIndex = i
        If due(i) >= monthStart And nextIndex = 0 Then nextIndex = i
    Next
    If lastIndex > 0 Then
        CounselingWindow due, lastIndex, windowStart, windowEnd
        recorded = MajorScalar("SELECT Count(*) FROM MemberCounseling WHERE MemberID=" & memberID & " AND IsArchived=False AND CounseledOn>=" & SQLDate(windowStart) & " AND CounseledOn<=" & SQLDate(windowEnd))
        If Nz(recorded, 0) = 0 And asOf <= windowEnd Then result("Status") = "Overdue": result("DueMonth") = due(lastIndex): Set CounselingMemberStatus = result: Exit Function
    End If
    If nextIndex = 0 Then result("Status") = "Not scheduled": Set CounselingMemberStatus = result: Exit Function
    result("DueMonth") = due(nextIndex)
    CounselingWindow due, nextIndex, windowStart, windowEnd
    recorded = MajorScalar("SELECT Count(*) FROM MemberCounseling WHERE MemberID=" & memberID & " AND IsArchived=False AND CounseledOn>=" & SQLDate(windowStart) & " AND CounseledOn<=" & SQLDate(asOf))
    If Nz(recorded, 0) > 0 Then
        result("Status") = "Recorded"
    ElseIf Month(due(nextIndex)) = Month(asOf) And Year(due(nextIndex)) = Year(asOf) Then
        result("Status") = "Due this month"
    ElseIf due(nextIndex) <= DateAdd("d", 60, asOf) Then
        result("Status") = "Due soon"
    Else
        result("Status") = "Scheduled"
    End If
    Set CounselingMemberStatus = result
End Function
Private Sub CounselingWindow(ByVal due As Collection, ByVal index As Long, ByRef windowStart As Date, ByRef windowEnd As Date)
    Dim monthStart As Date
    monthStart = DateSerial(Year(due(index)), Month(due(index)), 1)
    If index > 1 Then windowStart = DateAdd("d", Int((monthStart - due(index - 1)) / 2) + 1, due(index - 1)) Else windowStart = DateAdd("m", -6, monthStart)
    If index < due.Count Then windowEnd = DateAdd("d", Int((DateSerial(Year(due(index + 1)), Month(due(index + 1)), 1) - due(index)) / 2), due(index)) Else windowEnd = DateAdd("m", 6, due(index))
End Sub
Private Function SQLDate(ByVal value As Date) As String
    SQLDate = "#" & Format$(value, "yyyy-mm-dd") & "#"
End Function

'A worksheet disposition is due once the member has a saved Detachment of
'Individual report and a completed session still has no recorded disposition.
Public Function CounselingDispositionDue(ByVal memberID As Long) As Boolean
    If Nz(MajorScalar("SELECT Count(*) FROM MemberCounseling WHERE MemberID=" & memberID & " AND IsArchived=False AND Outcome='" & COUNSELING_COMPLETED & "' AND WorksheetDisposition Is Null"), 0) = 0 Then Exit Function
    CounselingDispositionDue = Nz(MajorScalar("SELECT Count(*) FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & memberID & " AND r.IsArchived=False AND r.OccasionDetachment='Yes'"), 0) > 0
End Function

Public Function SaveCounseling(ByVal counselingID As Long, ByVal expectedVersion As Long, ByVal memberID As Long, ByVal counseledOn As Date, ByVal outcome As String, ByVal counselor As String, ByVal reason As String, ByVal disposition As String, ByVal dispositionOn As Variant) As Long
    Dim r As DAO.Recordset, w As DAO.Workspace, author As Long, op As Long
    author = MajorAuthor()
    counselor = Trim$(counselor): reason = Trim$(reason): disposition = Trim$(disposition)
    If outcome <> COUNSELING_COMPLETED And outcome <> COUNSELING_NOT_PERFORMED Then Err.Raise vbObjectError + 9500, "modCounseling.SaveCounseling", "Choose Completed or Not performed."
    If outcome = COUNSELING_COMPLETED And (counselor = "" Or Len(counselor) > 20) Then Err.Raise vbObjectError + 9501, "modCounseling.SaveCounseling", "Enter the counselor as it appears in block 31 (20 characters or fewer)."
    If outcome = COUNSELING_NOT_PERFORMED And (reason = "" Or Len(reason) > 20) Then Err.Raise vbObjectError + 9502, "modCounseling.SaveCounseling", "Enter a brief reason such as TEMADD or ILLNESS (20 characters or fewer)."
    If outcome = COUNSELING_NOT_PERFORMED Then counselor = ""
    If outcome = COUNSELING_COMPLETED Then reason = ""
    If disposition <> "" And disposition <> "Given to member" And disposition <> "Destroyed" Then Err.Raise vbObjectError + 9503, "modCounseling.SaveCounseling", "Choose Given to member or Destroyed for the worksheet."
    'The disposition date is recorded automatically; nobody needs to enter it.
    If disposition <> "" And Not IsDate(dispositionOn) Then dispositionOn = Date
    If disposition = "" Then dispositionOn = Null
    If DateValue(counseledOn) > Date Then Err.Raise vbObjectError + 9504, "modCounseling.SaveCounseling", "The counseling date cannot be in the future."
    If Nz(MajorScalar("SELECT IsArchived FROM CommandMember WHERE MemberID=" & memberID), True) Then Err.Raise vbObjectError + 9505, "modCounseling.SaveCounseling", "Choose an active command member."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM MemberCounseling WHERE CounselingID=" & counselingID, dbOpenDynaset)
    If counselingID = 0 Then
        r.AddNew: r!CounselingGuid = NewIdentity(): r!CreatedBy = author: r!CreatedAt = Now(): r!VersionCounter = 1: r!IsArchived = False
    Else
        If r.EOF Then Err.Raise vbObjectError + 9506, "modCounseling.SaveCounseling", "The counseling record no longer exists."
        If r!MemberID <> memberID Or r!VersionCounter <> expectedVersion Then Err.Raise vbObjectError + 9507, "modCounseling.SaveCounseling", "The counseling record changed. Reopen it before editing."
        If r!IsArchived Then Err.Raise vbObjectError + 9508, "modCounseling.SaveCounseling", "Archived counseling records are read only."
        r.Edit: r!VersionCounter = r!VersionCounter + 1
    End If
    r!MemberID = memberID: r!CounseledOn = DateValue(counseledOn): r!Outcome = outcome
    r!CounselorName = IIf(counselor = "", Null, counselor): r!NotPerformedReason = IIf(reason = "", Null, reason)
    r!WorksheetDisposition = IIf(disposition = "", Null, disposition): r!DispositionOn = IIf(IsNull(dispositionOn), Null, DateValue(Nz(dispositionOn, Date)))
    r!ModifiedBy = author: r!ModifiedAt = Now()
    r.Update: r.Bookmark = r.LastModified: SaveCounseling = r!CounselingID: r.Close
    op = BeginOperation("Save counseling record", "Member " & memberID & "; counseling record " & SaveCounseling & "; " & outcome & " " & Format$(counseledOn, "yyyy-mm-dd"))
    FinishOperation op, "Completed", "Counseling record saved. Reports are unchanged until a contributor applies it."
    w.CommitTrans: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modCounseling", "SaveCounseling", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    SwiftRollback w, swiftPrimary: If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function
'Common block 31 reasons; any other brief reason may be typed.
Public Function CounselingReasonChoices() As String
    CounselingReasonChoices = """TEMADD"";""ILLNESS"";""LEAVE"";""DEPLOYED"";""HOSPITALIZED"";""SEE COMMENT"""
End Function
Public Function ArchiveCounseling(ByVal counselingID As Long, ByVal expectedVersion As Long) As Boolean
    Dim r As DAO.Recordset, w As DAO.Workspace, op As Long
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = CurrentDb.OpenRecordset("SELECT * FROM MemberCounseling WHERE CounselingID=" & counselingID, dbOpenDynaset)
    If r.EOF Then Err.Raise vbObjectError + 9506, "modCounseling.ArchiveCounseling", "The counseling record no longer exists."
    If r!VersionCounter <> expectedVersion Then Err.Raise vbObjectError + 9507, "modCounseling.ArchiveCounseling", "The counseling record changed. Reopen it before archiving."
    If r!IsArchived Then Err.Raise vbObjectError + 9508, "modCounseling.ArchiveCounseling", "The counseling record is already archived."
    r.Edit: r!IsArchived = True: r!VersionCounter = r!VersionCounter + 1: r!ModifiedBy = MajorAuthor(): r!ModifiedAt = Now(): r.Update: r.Close
    op = BeginOperation("Archive counseling record", "Counseling record " & counselingID)
    FinishOperation op, "Completed", "Counseling record archived; it is retained and no longer counts toward the schedule."
    w.CommitTrans: ArchiveCounseling = True: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modCounseling", "ArchiveCounseling", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    SwiftRollback w, swiftPrimary: If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function

'My Work rows: member ID;member;rank;status;due month;last counseled;note. Only
'members needing attention are listed unless everything is requested.
Public Function CounselingWorkRows(Optional ByVal includeAll As Boolean = False, Optional ByVal asOf As Date = 0) As String
    Dim r As DAO.Recordset, s As Object, rows As String, status As String
    Set r = CurrentDb.OpenRecordset("SELECT MemberID,FullName,RankRate FROM CommandMember WHERE IsArchived=False ORDER BY FullName", dbOpenSnapshot)
    Do While Not r.EOF
        Set s = CounselingMemberStatus(r!MemberID, asOf): status = s("Status")
        If CounselingDispositionDue(r!MemberID) Then status = "Worksheet disposition due"
        If includeAll Or status = "Overdue" Or status = "Due this month" Or status = "Due soon" Or status = "Worksheet disposition due" Then
            If rows <> "" Then rows = rows & ";"
            rows = rows & r!MemberID & ";" & ListText(r!FullName) & ";" & ListText(Nz(r!RankRate, "")) & ";" & status & ";" & IIf(IsNull(s("DueMonth")), "", Format$(Nz(s("DueMonth"), Date), "yyyy-mm")) & ";" & IIf(IsNull(s("LastCounseled")), "", Format$(Nz(s("LastCounseled"), Date), "yyyy-mm-dd")) & ";" & ListText(s("Note"))
        End If
        r.MoveNext
    Loop
    r.Close
    CounselingWorkRows = rows
End Function
Private Function ListText(ByVal value As Variant) As String
    ListText = """" & Replace(Replace(Nz(value, ""), ";", ","), """", "'") & """"
End Function

'Suggested blocks 30-31 for a Regular report from the member's records. Returns
'Nothing when no suggestion applies or the report already matches it. A completed
'session already entered on another of the member's reports is never offered twice.
Public Function CounselingSuggestion(ByVal f As Object, ByVal reportID As Long) As Object
    Dim memberID As Long, s As Object, r As DAO.Recordset, fromDate As Date, toDate As Date, dateText As String, existing As String, records As Long
    If reportID = 0 Then Exit Function
    If Not RequiresCounselingEntry(f) Then Exit Function
    'Midterm counseling is recorded on the next Regular report (blocks 30-32 guidance).
    If FieldText(f, "ReportRegular") <> "Yes" Then Exit Function
    If Not IsDate(FieldText(f, "PeriodStart")) Or Not IsDate(FieldText(f, "PeriodEnd")) Then Exit Function
    memberID = Nz(MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & reportID), 0)
    If memberID = 0 Then Exit Function
    records = Nz(MajorScalar("SELECT Count(*) FROM MemberCounseling WHERE MemberID=" & memberID & " AND IsArchived=False"), 0)
    If records = 0 Then Exit Function
    fromDate = CDate(FieldText(f, "PeriodStart")): toDate = CDate(FieldText(f, "PeriodEnd"))
    Set s = CounselingMatch(memberID, fromDate, toDate, reportID)
    If s Is Nothing Then
        'A session in the period that another report already uses is not "no counseling".
        If Nz(MajorScalar("SELECT Count(*) FROM MemberCounseling WHERE MemberID=" & memberID & " AND IsArchived=False AND CounseledOn>=" & SQLDate(fromDate) & " AND CounseledOn<=" & SQLDate(toDate)), 0) > 0 Then Exit Function
        Set s = CreateObject("Scripting.Dictionary")
        s("CounselingID") = 0: s("CounselingDate") = "NOT REQ": s("CounselorName") = ""
        s("Message") = "No counseling is recorded for this member in this report period."
    End If
    existing = FieldText(f, "CounselingDate")
    If existing = s("CounselingDate") And FieldText(f, "CounselorName") = Nz(s("CounselorName"), "") Then Exit Function
    'NOT REQ is only offered for an empty block 30; any entry there is deliberate.
    If s("CounselingDate") = "NOT REQ" And existing <> "" Then Exit Function
    Set CounselingSuggestion = s
End Function
'The member's recorded session for a report period: the latest completed session
'not already entered on another of the member's saved reports, else a session
'recorded as not performed. Nothing when neither exists.
Private Function CounselingMatch(ByVal memberID As Long, ByVal fromDate As Date, ByVal toDate As Date, ByVal reportID As Long) As Object
    Dim s As Object, r As DAO.Recordset, dateText As String
    Set s = CreateObject("Scripting.Dictionary")
    Set r = CurrentDb.OpenRecordset("SELECT * FROM MemberCounseling WHERE MemberID=" & memberID & " AND IsArchived=False AND CounseledOn>=" & SQLDate(fromDate) & " AND CounseledOn<=" & SQLDate(toDate) & " ORDER BY CounseledOn DESC", dbOpenSnapshot)
    Do While Not r.EOF
        dateText = Format$(r!CounseledOn, "yyyy-mm-dd")
        If r!Outcome = COUNSELING_COMPLETED Then
            'Only another Regular report uses up the session; a Concurrent-only report
            'repeating the date does not.
            If Nz(MajorScalar("SELECT Count(*) FROM Report AS p INNER JOIN MemberReport AS m ON p.ReportID=m.ReportID WHERE m.MemberID=" & memberID & " AND p.ReportID<>" & reportID & " AND p.IsArchived=False AND p.CounselingDate=" & QuoteSQL(dateText) & _
                " AND p.CurrentRevisionID Not In (SELECT c.RevisionID FROM RevisionValue AS c WHERE c.FieldKey='ReportConcurrent' AND c.ValueText='Yes' AND c.RevisionID Not In (SELECT g.RevisionID FROM RevisionValue AS g WHERE g.FieldKey='ReportRegular' AND g.ValueText='Yes'))"), 0) = 0 Then
                s("CounselingID") = r!CounselingID: s("CounselingDate") = dateText: s("CounselorName") = r!CounselorName
                s("Message") = "Recorded counseling on " & dateText & " (" & r!CounselorName & ") falls in this report period."
                Exit Do
            End If
        ElseIf s.Count = 0 Then
            s("CounselingID") = r!CounselingID: s("CounselingDate") = "NOT PERF": s("CounselorName") = r!NotPerformedReason
            s("Message") = "Counseling was recorded as not performed (" & r!NotPerformedReason & ") in this report period."
        End If
        r.MoveNext
    Loop
    r.Close
    If s.Count > 0 Then Set CounselingMatch = s
End Function
'New reports for a member start with blocks 30-31 from the recorded session in
'their period. Entered values are never replaced, and Concurrent-only or
'Operational Commander reports are left blank.
Public Sub CounselingCreationDefaults(ByVal values As Object, ByVal memberID As Long)
    Dim s As Object
    If memberID = 0 Then Exit Sub
    If Not IsDate(FieldText(values, "PeriodStart")) Or Not IsDate(FieldText(values, "PeriodEnd")) Then Exit Sub
    If FieldText(values, "CounselingDate") <> "" Or FieldText(values, "CounselorName") <> "" Then Exit Sub
    If FieldText(values, "ReportOperationalCommander") = "Yes" Then Exit Sub
    If FieldText(values, "ReportConcurrent") = "Yes" And FieldText(values, "ReportRegular") <> "Yes" Then Exit Sub
    Set s = CounselingMatch(memberID, CDate(FieldText(values, "PeriodStart")), CDate(FieldText(values, "PeriodEnd")), 0)
    If s Is Nothing Then Exit Sub
    values("CounselingDate") = s("CounselingDate"): values("CounselorName") = s("CounselorName")
End Sub
Private Function FieldText(ByVal f As Object, ByVal key As String) As String
    If f.Exists(key) Then FieldText = Trim$(Nz(f(key), ""))
End Function
Public Sub CounselingFindings(ByVal findings As Collection, ByVal f As Object, ByVal reportID As Long)
    Dim s As Object
    Set s = CounselingSuggestion(f, reportID)
    If s Is Nothing Then Exit Sub
    VAdd findings, f, reportID, "CounselingDate", "Needs review", "COUNSEL-RECORD", s("Message") & " Use recorded counseling enters " & s("CounselingDate") & IIf(Nz(s("CounselorName"), "") = "", "", " / " & s("CounselorName")) & " in blocks 30-31."
End Sub
