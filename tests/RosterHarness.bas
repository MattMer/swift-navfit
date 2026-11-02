Attribute VB_Name = "RosterHarness"
Option Explicit
'F02: roster snapshots, departures from PRD and from absence, matching and rank.
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 7790, , message
    checks = checks + 1
End Sub
'One synthetic NRRM roster; each person is "NAME|RATE|UIC|PRD".
Private Function WriteRoster(ByVal people As Variant) As String
    Dim stream As Object, text As String, key As Variant, person As Variant, parts As Variant, cells As String, value As String
    For Each key In RUADKeys(): text = text & IIf(Len(text) > 0, ",", "") & key: Next
    For Each person In people
        parts = Split(person, "|"): cells = ""
        For Each key In RUADKeys()
            Select Case key
                Case "NAME": value = parts(0)
                Case "DESG", "RATE/RANK": value = parts(1)
                Case "UIC", "TRUIC", "UMUIC": value = parts(2)
                Case "PRD": value = parts(3)
                Case "E/O": value = "Enlisted"
                Case "SELRES TYPE": value = "Drilling"
                Case Else: value = "SYN"
            End Select
            cells = cells & IIf(Len(cells) > 0, ",", "") & value
        Next
        text = text & vbCrLf & cells
    Next
    WriteRoster = Environ$("TEMP") & "\Swift-roster-" & Mid$(NewIdentity(), 2, 36) & ".csv"
    Set stream = CreateObject("ADODB.Stream"): stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText text: stream.SaveToFile WriteRoster, 2: stream.Close
End Function
'Load, match every row (refresh the named member or create one), preview, apply.
Private Function ImportRoster(ByVal people As Variant, ByVal complete As Boolean, Optional ByVal rankFor As String = "") As Long
    Dim path As String, i As Long, row As Object, summary As String
    path = WriteRoster(people)
    Call RUADLoad(path): Call RUADSuggest
    For i = 1 To RUADCount()
        Set row = RUADRow(i)
        If row("Suggested") > 0 Then
            RUADChoose i, "Refresh existing", CLng(row("Suggested")), False, (CStr(row("Data")("NAME")) = rankFor)
        Else
            RUADChoose i, "Create new", 0
        End If
    Next
    RUADSetComplete complete
    summary = RUADPreview()
    Call RUADApply
    Kill path
    ImportRoster = DMax("RosterImportID", "RosterImport")
End Function
Private Function Departures(ByVal name As String, ByVal state As String, Optional ByVal basis As String = "") As Long
    Departures = Nz(MajorScalar("SELECT Count(*) FROM MemberDeparture AS d INNER JOIN CommandMember AS c ON d.MemberID=c.MemberID WHERE c.FullName='" & name & "' AND d.DepartureState='" & state & "'" & IIf(Len(basis) > 0, " AND d.Basis='" & basis & "'", "")), 0)
End Function
Private Function ResolveRefused(ByVal id As Long, ByVal outcome As String, ByVal detach As Variant, ByVal reason As String) As Boolean
    On Error GoTo Refused
    ResolveDeparture id, DLookup("VersionCounter", "MemberDeparture", "DepartureID=" & id), outcome, detach, reason
    Exit Function
Refused:
    ResolveRefused = True
End Function
Public Function TestRoster() As String
    Dim soon As String, later As String, first As Long, id As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC ROSTER", "Coordinator"
    ' PRD formats NRRM and MNA exports use; a month-only PRD means month end.
    Check RosterPRDDate("2027-03") = #3/31/2027# And RosterPRDDate("202703") = #3/31/2027# And RosterPRDDate("0327") = #3/31/2027#, "Month PRD not read"
    Check RosterPRDDate("03/2027") = #3/31/2027# And RosterPRDDate("2027-03-15") = #3/15/2027# And RosterPRDDate("03/15/2027") = #3/15/2027# And RosterPRDDate("20270315") = #3/15/2027#, "Dated PRD not read"
    Check IsNull(RosterPRDDate("SYN")) And IsNull(RosterPRDDate("")) And IsNull(RosterPRDDate("2027-13")) And IsNull(RosterPRDDate("02/30/2027")), "Unreadable PRD accepted"

    soon = Format$(DateAdd("d", 45, Date), "yyyy-mm-dd"): later = Format$(DateAdd("yyyy", 2, Date), "yyyy-mm")
    ' First complete roster: three people, one with a PRD inside the window.
    first = ImportRoster(Array("ALPHA ONE|IT2|11111|" & later, "BRAVO TWO|IT2|11111|" & later, "CHARLIE THREE|IT2|11111|" & soon), True)
    Check DCount("*", "RosterPresence", "RosterImportID=" & first) = 3 And DLookup("ScopeUICs", "RosterImport", "RosterImportID=" & first) = "|11111|", "First snapshot wrong"
    Check Departures("CHARLIE THREE", "Pending", "PRD") = 1 And Departures("ALPHA ONE", "Pending") = 0, "PRD inside the window not flagged"
    Check DLookup("ExpectedDate", "MemberDeparture", "Basis='PRD'") = DateValue(soon), "PRD departure date wrong"

    ' Second complete roster without BRAVO: flagged absent; matching refreshed the others.
    Call ImportRoster(Array("ALPHA ONE|IT2|11111|" & later, "CHARLIE THREE|IT2|11111|" & soon), True)
    Check DCount("*", "CommandMember") = 3, "Matching created duplicate profiles"
    Check Departures("BRAVO TWO", "Pending", "Absent from roster") = 1 And Departures("CHARLIE THREE", "Pending") = 1, "Absent member not flagged, or PRD departure duplicated"

    ' A partial roster flags no one as missing.
    Call ImportRoster(Array("ALPHA ONE|IT2|11111|" & later), False)
    Check Departures("CHARLIE THREE", "Pending", "Absent from roster") = 0 And DCount("*", "MemberDeparture") = 2, "Partial roster flagged missing members"

    ' BRAVO back on a complete roster: Returned. A promotion updates rank only when chosen.
    Call ImportRoster(Array("ALPHA ONE|IT1|11111|" & later, "BRAVO TWO|IT1|11111|" & later, "CHARLIE THREE|IT2|11111|" & soon), True, "BRAVO TWO")
    Check Departures("BRAVO TWO", "Returned") = 1 And Departures("BRAVO TWO", "Pending") = 0, "Returned member still pending"
    Check DLookup("RankRate", "CommandMember", "FullName='BRAVO TWO'") = "IT1" And DLookup("RankRate", "CommandMember", "FullName='ALPHA ONE'") = "IT2", "Rank update ignored the coordinator's choice"

    ' Confirm needs a date; Dismiss needs a reason; both keep the history.
    id = DLookup("DepartureID", "MemberDeparture", "Basis='PRD' AND DepartureState='Pending'")
    Check ResolveRefused(id, "Confirm", Null, "") And ResolveRefused(id, "Dismiss", Null, ""), "Resolve accepted missing date or reason"
    Check PendingDepartureCount() = 1 And InStr(DepartureRows(), "CHARLIE THREE") > 0, "Departure list wrong"
    ResolveDeparture id, DLookup("VersionCounter", "MemberDeparture", "DepartureID=" & id), "Confirm", DateValue(soon), "Orders received"
    Check DLookup("DepartureState", "MemberDeparture", "DepartureID=" & id) = "Confirmed" And DLookup("ConfirmedDate", "MemberDeparture", "DepartureID=" & id) = DateValue(soon), "Confirm not recorded"
    Check PendingDepartureCount() = 0 And InStr(DepartureRows(), "CHARLIE THREE") > 0 And InStr(DepartureRows(), "BRAVO TWO") = 0 And InStr(DepartureRows(True), "BRAVO TWO") > 0, "Open and all-history lists wrong"
    Check ResolveRefused(id, "Dismiss", Null, "Changed mind"), "A confirmed departure was dismissed"
    TestRoster = "PASS: " & checks & " roster snapshot, departure, matching and rank assertions": Exit Function
Failed:
    TestRoster = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
