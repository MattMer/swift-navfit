Attribute VB_Name = "modRoster"
Option Compare Binary
Option Explicit
'F02: each roster import is a dated snapshot of who was on it. A member missing
'from a newer complete roster for their UIC, or whose PRD is inside the lead
'window, becomes a pending departure for the coordinator to confirm or dismiss.
'Nothing is archived or deleted automatically.
Public Const DEPARTURE_LEAD_DAYS As Long = 90

'A roster PRD as a date: yyyy-mm-dd, mm/dd/yyyy, yyyymmdd, yyyy-mm, yyyymm,
'mm/yyyy or mmyy. A month-only PRD means the last day of that month. Null when
'the text cannot be read.
Public Function RosterPRDDate(ByVal text As Variant) As Variant
    Dim t As String, parts As Variant, y As Long, m As Long, d As Long
    RosterPRDDate = Null
    t = Trim$(Nz(text, ""))
    If Len(t) = 0 Then Exit Function
    t = Replace(Replace(t, ".", "-"), "/", "-")
    If t Like "####-##-##" Then
        y = Val(Left$(t, 4)): m = Val(Mid$(t, 6, 2)): d = Val(Right$(t, 2))
    ElseIf t Like "##-##-####" Or t Like "#-##-####" Or t Like "##-#-####" Or t Like "#-#-####" Then
        parts = Split(t, "-"): m = Val(parts(0)): d = Val(parts(1)): y = Val(parts(2))
    ElseIf t Like "########" Then
        y = Val(Left$(t, 4)): m = Val(Mid$(t, 5, 2)): d = Val(Right$(t, 2))
    ElseIf t Like "####-##" Or t Like "####-#" Then
        parts = Split(t, "-"): y = Val(parts(0)): m = Val(parts(1))
    ElseIf t Like "######" Then
        y = Val(Left$(t, 4)): m = Val(Right$(t, 2))
    ElseIf t Like "##-####" Or t Like "#-####" Then
        parts = Split(t, "-"): m = Val(parts(0)): y = Val(parts(1))
    ElseIf t Like "####" Then
        m = Val(Left$(t, 2)): y = 2000 + Val(Right$(t, 2))
    Else
        Exit Function
    End If
    If y < 1900 Or y > 2200 Or m < 1 Or m > 12 Then Exit Function
    If d = 0 Then
        RosterPRDDate = DateSerial(y, m + 1, 0)
    Else
        If d > Day(DateSerial(y, m + 1, 0)) Then Exit Function
        RosterPRDDate = DateSerial(y, m, d)
    End If
End Function

'Record one roster import and who was on it, then flag departures. Runs inside
'the caller's import transaction. present maps MemberID -> Array(UIC, PRD, rank).
Public Function RecordRosterImport(ByVal db As DAO.Database, ByVal sourceFormat As String, ByVal sourceName As String, ByVal sourceHash As String, ByVal scopeUICs As String, ByVal complete As Boolean, ByVal rosterRows As Long, ByVal present As Object) As Long
    Dim r As DAO.Recordset, id As Long, key As Variant, item As Variant
    Set r = db.OpenRecordset("RosterImport", dbOpenDynaset, dbAppendOnly)
    r.AddNew: r!ImportGuid = NewIdentity(): r!SourceFormat = sourceFormat: r!SourceName = Left$(sourceName, 255): r!SourceHash = sourceHash
    r!AsOfDate = Date: r!ScopeUICs = scopeUICs: r!IsComplete = complete: r!RosterRows = rosterRows
    r!ImportedBy = MajorAuthor(): r!ImportedAt = Now()
    r.Update: r.Bookmark = r.LastModified: id = r!RosterImportID: r.Close
    Set r = db.OpenRecordset("RosterPresence", dbOpenDynaset, dbAppendOnly)
    For Each key In present.Keys
        item = present(key)
        r.AddNew: r!RosterImportID = id: r!MemberID = CLng(key): r!UIC = Left$(item(0), 12): r!PRD = Left$(item(1), 20): r!RosterRank = Left$(item(2), 12): r.Update
    Next
    r.Close
    DetectDepartures db, id
    RecordRosterImport = id
End Function

Private Function OpenDeparture(ByVal db As DAO.Database, ByVal memberID As Long) As Variant
    OpenDeparture = Null
    Dim r As DAO.Recordset
    Set r = db.OpenRecordset("SELECT DepartureID FROM MemberDeparture WHERE MemberID=" & memberID & " AND DepartureState In ('Pending','Confirmed') ORDER BY DepartureID DESC", dbOpenSnapshot)
    If Not r.EOF Then OpenDeparture = r!DepartureID
    r.Close
End Function

Private Sub AddDeparture(ByVal db As DAO.Database, ByVal memberID As Long, ByVal basis As String, ByVal importID As Long, ByVal prd As Variant)
    Dim r As DAO.Recordset
    Set r = db.OpenRecordset("MemberDeparture", dbOpenDynaset, dbAppendOnly)
    r.AddNew: r!MemberID = memberID: r!Basis = basis: r!RosterImportID = importID: r!PRD = prd: r!ExpectedDate = prd
    r!DepartureState = "Pending": r!CreatedBy = MajorAuthor(): r!CreatedAt = Now(): r!VersionCounter = 1
    r.Update: r.Close
End Sub

'Absent from a complete roster: present in the last earlier complete roster for the
'same UIC, active, and missing now. Back on the roster: an open absence is marked
'Returned. PRD: a readable PRD inside the lead window opens (or updates) a PRD
'departure.
Private Sub DetectDepartures(ByVal db As DAO.Database, ByVal importID As Long)
    Dim snapshot As DAO.Recordset, r As DAO.Recordset, complete As Boolean, scope As String, uic As Variant, previous As Variant, prd As Variant, openID As Variant
    Set snapshot = db.OpenRecordset("SELECT * FROM RosterImport WHERE RosterImportID=" & importID, dbOpenSnapshot)
    complete = Nz(snapshot!IsComplete, False): scope = Nz(snapshot!ScopeUICs, ""): snapshot.Close
    db.Execute "UPDATE MemberDeparture SET DepartureState='Returned',ResolvedBy=" & MajorAuthor() & ",ResolvedAt=Now(),Reason='Back on the roster imported " & Format$(Date, "yyyy-mm-dd") & "',VersionCounter=VersionCounter+1 WHERE DepartureState='Pending' AND Basis='Absent from roster' AND MemberID In (SELECT MemberID FROM RosterPresence WHERE RosterImportID=" & importID & ")", dbFailOnError
    If complete Then
        For Each uic In Split(Mid$(scope, 2), "|")
            If Len(uic) > 0 Then
                previous = MajorScalar("SELECT Max(RosterImportID) FROM RosterImport WHERE IsComplete=True AND RosterImportID<" & importID & " AND ScopeUICs Like " & QuoteSQL("*|" & uic & "|*"))
                If Not IsNull(previous) Then
                    Set r = db.OpenRecordset("SELECT p.MemberID FROM RosterPresence AS p INNER JOIN CommandMember AS c ON p.MemberID=c.MemberID WHERE p.RosterImportID=" & previous & " AND p.UIC=" & QuoteSQL(CStr(uic)) & " AND c.IsArchived=False AND p.MemberID Not In (SELECT MemberID FROM RosterPresence WHERE RosterImportID=" & importID & ")", dbOpenSnapshot)
                    Do While Not r.EOF
                        If IsNull(OpenDeparture(db, r!MemberID)) Then AddDeparture db, r!MemberID, "Absent from roster", importID, Null
                        r.MoveNext
                    Loop
                    r.Close
                End If
            End If
        Next
    End If
    Set r = db.OpenRecordset("SELECT MemberID,PRD FROM RosterPresence WHERE RosterImportID=" & importID, dbOpenSnapshot)
    Do While Not r.EOF
        prd = RosterPRDDate(r!PRD)
        If Not IsNull(prd) Then
            If prd <= DateAdd("d", DEPARTURE_LEAD_DAYS, Date) Then
                openID = OpenDeparture(db, r!MemberID)
                If IsNull(openID) Then
                    AddDeparture db, r!MemberID, "PRD", importID, prd
                Else
                    db.Execute "UPDATE MemberDeparture SET PRD=" & SQLDay(prd) & ",ExpectedDate=" & SQLDay(prd) & ",VersionCounter=VersionCounter+1 WHERE DepartureID=" & openID & " AND Basis='PRD' AND DepartureState='Pending' AND (PRD Is Null OR PRD<>" & SQLDay(prd) & ")", dbFailOnError
                End If
            End If
        End If
        r.MoveNext
    Loop
    r.Close
End Sub
Private Function SQLDay(ByVal value As Date) As String
    SQLDay = "#" & Format$(value, "yyyy-mm-dd") & "#"
End Function

'Confirm records the actual detachment date; Dismiss needs a reason (for example,
'still aboard and missing from that export). Both keep the history.
Public Sub ResolveDeparture(ByVal departureID As Long, ByVal expectedVersion As Long, ByVal outcome As String, ByVal detachDate As Variant, ByVal reason As String)
    Dim r As DAO.Recordset, operation As Long
    If outcome <> "Confirm" And outcome <> "Dismiss" Then Err.Raise vbObjectError + 9701, "modRoster.ResolveDeparture", "Choose Confirm or Dismiss."
    If outcome = "Confirm" And Not IsDate(detachDate) Then Err.Raise vbObjectError + 9702, "modRoster.ResolveDeparture", "Enter the detachment date to confirm the departure."
    If outcome = "Dismiss" And Len(Trim$(reason)) = 0 Then Err.Raise vbObjectError + 9703, "modRoster.ResolveDeparture", "Say why the member is not departing."
    Set r = CurrentDb.OpenRecordset("SELECT * FROM MemberDeparture WHERE DepartureID=" & departureID, dbOpenDynaset)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 9704, "modRoster.ResolveDeparture", "Departure no longer exists."
    If r!VersionCounter <> expectedVersion Then r.Close: Err.Raise vbObjectError + 9705, "modRoster.ResolveDeparture", "This departure changed. Refresh and try again."
    If r!DepartureState <> "Pending" And Not (r!DepartureState = "Confirmed" And outcome = "Confirm") Then r.Close: Err.Raise vbObjectError + 9706, "modRoster.ResolveDeparture", "Only a pending departure can be resolved."
    operation = BeginOperation(outcome & " departure", "Departure " & departureID & ", member " & r!MemberID)
    r.Edit
    If outcome = "Confirm" Then
        r!DepartureState = "Confirmed": r!ConfirmedDate = DateValue(detachDate): r!ExpectedDate = DateValue(detachDate)
    Else
        r!DepartureState = "Dismissed"
    End If
    If Len(Trim$(reason)) > 0 Then r!Reason = Left$(reason, 255)
    r!ResolvedBy = MajorAuthor(): r!ResolvedAt = Now(): r!VersionCounter = r!VersionCounter + 1
    r.Update: r.Close
    FinishOperation operation, "Completed", "Departure " & LCase$(outcome) & "ed; the member profile is unchanged."
End Sub

'My Work list rows: id; member; rate; basis; expected date; state; detected.
Public Function DepartureRows(Optional ByVal includeAll As Boolean = False) As String
    Dim r As DAO.Recordset, rows As String, sql As String
    sql = "SELECT d.DepartureID,c.FullName,c.RankRate,d.Basis,d.ExpectedDate,d.DepartureState,d.CreatedAt FROM MemberDeparture AS d INNER JOIN CommandMember AS c ON d.MemberID=c.MemberID WHERE c.IsArchived=False"
    If Not includeAll Then sql = sql & " AND d.DepartureState In ('Pending','Confirmed')"
    Set r = CurrentDb.OpenRecordset(sql & " ORDER BY d.ExpectedDate,c.FullName", dbOpenSnapshot)
    Do While Not r.EOF
        If rows <> "" Then rows = rows & ";"
        rows = rows & r!DepartureID & ";" & UXListCell(r!FullName) & ";" & UXListCell(Nz(r!RankRate, "")) & ";" & UXListCell(r!Basis) & ";" & UXListCell(IIf(IsNull(r!ExpectedDate), "Date unknown", Format$(Nz(r!ExpectedDate, Date), "yyyy-mm-dd"))) & ";" & UXListCell(r!DepartureState) & ";" & UXListCell(Format$(r!CreatedAt, "yyyy-mm-dd"))
        r.MoveNext
    Loop
    r.Close
    DepartureRows = rows
End Function
Public Function PendingDepartureCount() As Long
    PendingDepartureCount = Nz(MajorScalar("SELECT Count(*) FROM MemberDeparture AS d INNER JOIN CommandMember AS c ON d.MemberID=c.MemberID WHERE c.IsArchived=False AND d.DepartureState='Pending'"), 0)
End Function
