Attribute VB_Name = "modReportsDue"
Option Compare Binary
Option Explicit
'F03: reports the command owes, worked out on refresh from the member directory,
'the periodic calendar and F02 departures. Only dismissals are stored. A due
'report drops off once a matching draft exists, even unsaved to completion.
Public Const DUE_LEAD_DAYS As Long = 90
Public Const DUE_OVERDUE_DAYS As Long = 90

'The report form for a worn paygrade: EVAL E1-E6, CHIEFEVAL E7-E9, FITREP W and O.
Public Function DueFormKind(ByVal worn As String) As String
    Select Case Left$(UCase$(worn), 1)
        Case "E": If Val(Mid$(worn, 2)) >= 7 Then DueFormKind = "CHIEFEVAL" Else DueFormKind = "EVAL"
        Case "W", "O": DueFormKind = "FITREP"
    End Select
End Function

Private Function SQLDay(ByVal value As Date) As String
    SQLDay = "#" & Format$(value, "yyyy-mm-dd") & "#"
End Function

'A Regular report ending on the due date (periodic), or a Detachment of Individual
'report ending within 30 days of the departure (detachment), satisfies the item.
'Concurrent-only and archived reports do not.
Private Function DueSatisfied(ByVal memberID As Long, ByVal basis As String, ByVal due As Date) As Boolean
    Dim sql As String
    sql = "SELECT Count(*) FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & memberID & " AND r.IsArchived=False AND r.ReportRegular='Yes'"
    If basis = "Periodic" Then
        sql = sql & " AND r.PeriodEnd=" & SQLDay(due)
    Else
        sql = sql & " AND r.OccasionDetachment='Yes' AND r.PeriodEnd Between " & SQLDay(DateAdd("d", -30, due)) & " AND " & SQLDay(DateAdd("d", 30, due))
    End If
    DueSatisfied = Nz(MajorScalar(sql), 0) > 0
End Function
Private Function DueDismissed(ByVal memberID As Long, ByVal basis As String, ByVal due As Date) As Boolean
    DueDismissed = Nz(MajorScalar("SELECT Count(*) FROM ReportDueDismissal WHERE MemberID=" & memberID & " AND Basis=" & QuoteSQL(basis) & " AND DueDate=" & SQLDay(due) & " AND UndoneAt Is Null"), 0) > 0
End Function
Private Sub AddItem(ByVal items As Collection, ByVal r As DAO.Recordset, ByVal basis As String, ByVal kind As String, ByVal due As Date, ByVal asOf As Date)
    Dim status As String
    If due < asOf Then
        status = "Overdue"
    ElseIf DateDiff("d", asOf, due) <= 30 Then
        status = "Due in " & DateDiff("d", asOf, due) & " days"
    Else
        status = "Due " & Format$(due, "mmm yyyy")
    End If
    items.Add Array(CLng(r!MemberID), CStr(r!FullName), Nz(r!RankRate, ""), basis, kind, due, status)
End Sub

'Every due report for active members, as Array(MemberID, name, rate, basis,
'form kind, due date, status), soonest first. missingRank returns how many
'active members have no rank/rate, so no due date can be worked out for them.
Public Function ReportsDueItems(Optional ByVal asOf As Date = 0, Optional ByRef missingRank As Long) As Collection
    Dim items As New Collection, r As DAO.Recordset, d As DAO.Recordset, worn As String, kind As String, y As Long, due As Variant, sorted As New Collection, item As Variant, i As Long, placed As Boolean
    If asOf = 0 Then asOf = Date
    missingRank = 0
    Set r = CurrentDb.OpenRecordset("SELECT MemberID,FullName,RankRate FROM CommandMember WHERE IsArchived=False ORDER BY FullName", dbOpenSnapshot)
    Do While Not r.EOF
        worn = ValidationWornPaygrade(Nz(r!RankRate, "")): kind = DueFormKind(worn)
        If Len(worn) = 0 Or Len(kind) = 0 Then
            missingRank = missingRank + 1
        ElseIf Not ((worn = "E9" Or worn = "W5") And asOf >= #6/3/2026#) Then
            'Since NAVADMIN 132/26 (3 Jun 2026) E-9 and CWO5 have no Regular periodic report.
            For y = Year(asOf) - 1 To Year(asOf) + 1
                due = PeriodicExpectedDate(worn, DateSerial(y, 12, 31), Nz(r!RankRate, ""))
                If Not IsNull(due) Then
                    If due >= DateAdd("d", -DUE_OVERDUE_DAYS, asOf) And due <= DateAdd("d", DUE_LEAD_DAYS, asOf) Then
                        If Not DueSatisfied(r!MemberID, "Periodic", due) And Not DueDismissed(r!MemberID, "Periodic", due) Then AddItem items, r, "Periodic", kind, due, asOf
                    End If
                End If
            Next
        End If
        Set d = CurrentDb.OpenRecordset("SELECT ExpectedDate FROM MemberDeparture WHERE MemberID=" & r!MemberID & " AND DepartureState In ('Pending','Confirmed') AND ExpectedDate Is Not Null", dbOpenSnapshot)
        Do While Not d.EOF
            due = d!ExpectedDate
            If Len(kind) > 0 And due >= DateAdd("d", -DUE_OVERDUE_DAYS, asOf) And due <= DateAdd("d", DUE_LEAD_DAYS, asOf) Then
                If Not DueSatisfied(r!MemberID, "Detachment", due) And Not DueDismissed(r!MemberID, "Detachment", due) Then AddItem items, r, "Detachment", kind, due, asOf
            End If
            d.MoveNext
        Loop
        d.Close
        r.MoveNext
    Loop
    r.Close
    For Each item In items
        placed = False
        For i = 1 To sorted.Count
            If item(5) < sorted(i)(5) Then sorted.Add item, Before:=i: placed = True: Exit For
        Next
        If Not placed Then sorted.Add item
    Next
    Set ReportsDueItems = sorted
End Function

'List rows: "member|basis|yyyy-mm-dd"; member; rate; report; due date; status.
Public Function ReportsDueRows(Optional ByRef count As Long, Optional ByRef missingRank As Long) As String
    Dim item As Variant, rows As String, items As Collection
    Set items = ReportsDueItems(, missingRank): count = items.Count
    For Each item In items
        If rows <> "" Then rows = rows & ";"
        rows = rows & UXListCell(item(0) & "|" & item(3) & "|" & Format$(item(5), "yyyy-mm-dd")) & ";" & UXListCell(item(1)) & ";" & UXListCell(item(2)) & ";" & UXListCell(item(4) & IIf(item(3) = "Detachment", " (detachment)", " (periodic)")) & ";" & UXListCell(Format$(item(5), "yyyy-mm-dd")) & ";" & UXListCell(item(6))
    Next
    ReportsDueRows = rows
End Function

Private Sub ParseDueKey(ByVal key As String, ByRef memberID As Long, ByRef basis As String, ByRef due As Date)
    Dim parts As Variant
    parts = Split(key, "|")
    If UBound(parts) <> 2 Then Err.Raise vbObjectError + 9801, "modReportsDue.ParseDueKey", "Choose a due report."
    memberID = CLng(parts(0)): basis = CStr(parts(1)): due = DateValue(parts(2))
End Sub

'Not required, with a reason (for example, a detachment report is not needed because
'the periodic report falls close by). Undo brings the item back.
Public Function DismissDue(ByVal key As String, ByVal reason As String) As Long
    Dim memberID As Long, basis As String, due As Date, r As DAO.Recordset, operation As Long
    ParseDueKey key, memberID, basis, due
    If Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 9802, "modReportsDue.DismissDue", "Say why this report is not required."
    operation = BeginOperation("Dismiss due report", "Member " & memberID & ", " & basis & " due " & Format$(due, "yyyy-mm-dd"))
    Set r = CurrentDb.OpenRecordset("ReportDueDismissal", dbOpenDynaset, dbAppendOnly)
    r.AddNew: r!MemberID = memberID: r!Basis = basis: r!DueDate = due: r!Reason = Left$(reason, 255): r!DismissedBy = MajorAuthor(): r!DismissedAt = Now()
    r.Update: r.Bookmark = r.LastModified: DismissDue = r!DismissalID: r.Close
    FinishOperation operation, "Completed", "Marked not required; no report changed."
End Function
Public Sub UndoDueDismissal(ByVal dismissalID As Long)
    CurrentDb.Execute "UPDATE ReportDueDismissal SET UndoneAt=Now(),UndoneBy=" & MajorAuthor() & " WHERE DismissalID=" & dismissalID & " AND UndoneAt Is Null", dbFailOnError
End Sub

'Create the drafts several due items ask for and put them in new summary groups,
'one per report type, paygrade, basis and due date (a group shares one ending
'date). All or nothing: a refused item rolls back every draft and group. Returns
'the number of drafts; groupIDs receives the new groups.
Public Function CreateDueBatch(ByVal keys As Collection, Optional ByVal groupIDs As Collection) As Long
    Dim plan As Object, order As New Collection, key As Variant, memberID As Long, basis As String, due As Date, r As DAO.Recordset
    Dim worn As String, kind As String, slot As String, members As Collection, w As DAO.Workspace, trans As Boolean, stage As String
    Dim g As Long, id As Long, title As String, suffix As Long, parts As Variant, op As Long, failure As Long, source As String, message As String, item As Variant
    If keys Is Nothing Then Err.Raise vbObjectError + 9806, "modReportsDue.CreateDueBatch", "Select the due reports to create."
    If keys.Count = 0 Then Err.Raise vbObjectError + 9806, "modReportsDue.CreateDueBatch", "Select the due reports to create."
    Set plan = CreateObject("Scripting.Dictionary")
    For Each key In keys
        ParseDueKey CStr(key), memberID, basis, due
        Set r = CurrentDb.OpenRecordset("SELECT FullName,RankRate,IsArchived FROM CommandMember WHERE MemberID=" & memberID, dbOpenSnapshot)
        If r.EOF Then r.Close: Err.Raise vbObjectError + 9803, "modReportsDue.CreateDueBatch", "A selected member no longer exists. Refresh Reports due."
        worn = ValidationWornPaygrade(Nz(r!RankRate, "")): kind = DueFormKind(worn)
        If Len(kind) = 0 Then message = r!FullName & ": enter the member's rank/rate first." Else message = ""
        If r!IsArchived Then message = r!FullName & ": recover the member before creating a report."
        If Len(message) = 0 And DueSatisfied(memberID, basis, due) Then message = r!FullName & ": a matching report already exists."
        r.Close
        If Len(message) > 0 Then Err.Raise vbObjectError + 9807, "modReportsDue.CreateDueBatch", message & " No drafts were created."
        slot = kind & "|" & worn & "|" & basis & "|" & Format$(due, "yyyy-mm-dd")
        If Not plan.Exists(slot) Then Set members = New Collection: plan.Add slot, members: order.Add slot
        plan(slot).Add CStr(key)
    Next
    On Error GoTo Failed
    stage = "Begin due batch": Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    op = BeginOperation("Create due reports and groups", keys.Count & " due reports in " & plan.Count & IIf(plan.Count = 1, " group", " groups"))
    For Each key In order
        parts = Split(CStr(key), "|")
        title = parts(1) & " " & parts(0) & " " & LCase$(parts(2)) & ", ending " & parts(3): suffix = 1
        Do While Nz(MajorScalar("SELECT Count(*) FROM SummaryGroup WHERE Title=" & QuoteSQL(title)), 0) > 0
            suffix = suffix + 1: title = parts(1) & " " & parts(0) & " " & LCase$(parts(2)) & ", ending " & parts(3) & " (" & suffix & ")"
        Loop
        stage = "Create group " & title: g = CreateSummaryGroup(title, CStr(parts(0)), CStr(parts(1)))
        If Not groupIDs Is Nothing Then groupIDs.Add g
        For Each item In plan(key)
            stage = "Create draft " & item: id = CreateDueDraft(CStr(item))
            stage = "Add draft to " & title: AddGroupMember g, id, CStr(parts(1))
            CreateDueBatch = CreateDueBatch + 1
        Next
    Next
    FinishOperation op, "Completed", CreateDueBatch & " drafts and their summary groups committed together."
    stage = "Commit due batch": w.CommitTrans: trans = False
    Exit Function
Failed:
    failure = Err.Number: source = Err.Source: message = Err.Description
    If trans Then w.Rollback
    Err.Raise failure, source, stage & ": " & message & " No drafts or groups were created."
End Function

'Create the draft a due item asks for, linked to the member: the report form for
'their paygrade, Regular, Periodic or Detachment of Individual, ending on the due
'date and starting the day after their latest earlier report when there is one.
Public Function CreateDueDraft(ByVal key As String) As Long
    Dim memberID As Long, basis As String, due As Date, r As DAO.Recordset, values As Object, field As Variant, kind As String, start As Variant, failure As Long, source As String, message As String
    ParseDueKey key, memberID, basis, due
    Set r = CurrentDb.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & memberID, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 9803, "modReportsDue.CreateDueDraft", "The member no longer exists."
    If r!IsArchived Then r.Close: Err.Raise vbObjectError + 9803, "modReportsDue.CreateDueDraft", "Recover the member before creating a report."
    kind = DueFormKind(ValidationWornPaygrade(Nz(r!RankRate, "")))
    If Len(kind) = 0 Then r.Close: Err.Raise vbObjectError + 9804, "modReportsDue.CreateDueDraft", "Enter the member's rank/rate first."
    If DueSatisfied(memberID, basis, due) Then r.Close: Err.Raise vbObjectError + 9805, "modReportsDue.CreateDueDraft", "A matching report already exists for this member."
    start = MajorScalar("SELECT Max(r.PeriodEnd) FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & memberID & " AND r.IsArchived=False AND r.PeriodEnd<" & SQLDay(due))
    If Not IsNull(start) Then start = DateAdd("d", 1, start)
    Set values = CreateObject("Scripting.Dictionary")
    For Each field In DraftFieldKeys(): values.Add CStr(field), Null: Next
    For Each field In UXSharedKeys(): values(CStr(field)) = UXEffectiveDefault(0, CStr(field)): Next
    values("FormKind") = kind: values("PeriodStart") = start: values("PeriodEnd") = due
    values("OfficialName") = IIf(Len(Nz(r!ReportName, "")) > 0, r!ReportName, Null): values("RankRate") = r!RankRate: values("MemberSSN") = r!MemberSSN
    values("ReportRegular") = "Yes"
    If basis = "Periodic" Then values("OccasionPeriodic") = "Yes" Else values("OccasionDetachment") = "Yes"
    CounselingCreationDefaults values, memberID
    MemberCreationDefaults values, memberID
    On Error GoTo Failed
    DirectoryCreationContext memberID
    CreateDueDraft = CommitDraft(0, kind, kind & " - " & r!FullName, r!FullName, start, due, "", IIf(basis = "Periodic", "Create periodic report due ", "Create detachment report due ") & Format$(due, "yyyy-mm-dd"), values("OfficialComments"), _
        FieldsSubset(values, AdministrationKeys()), FieldsSubset(values, TraitKeys()), FieldsSubset(values, AssessmentKeys()), values("PromotionStatus"), FieldsSubset(values, ReportingSeniorKeys()), _
        FieldsSubset(values, ReportContextKeys()), FieldsSubset(values, MemberDetailsKeys()), FieldsSubset(values, ServiceDetailsKeys()), FieldsSubset(values, DecisionDetailsKeys()), FieldsSubset(values, ReportPeopleKeys()), FieldsSubset(values, PartyNamesKeys()), values("CommandEmployment"))
    DirectoryCreationContext 0
    r.Close
    Exit Function
Failed:
    failure = Err.Number: source = Err.Source: message = Err.Description
    DirectoryCreationContext 0
    Err.Raise failure, source, message
End Function
