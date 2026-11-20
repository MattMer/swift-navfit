Attribute VB_Name = "modSessionPresence"
Option Compare Database
Option Explicit
' Who else has this workspace open: up to 10 people can share one data file on a
' network drive. Each session keeps one SessionPresence row: its contributor, computer,
' when it started, when it was last seen, and the report open in its editor. A hidden
' keeper form (frmSessionKeeper) refreshes the row every minute while Swift NAVFIT is
' open and removes it when Access closes, so "connected now" means seen in the last few
' minutes. Old rows are cleared. Everything here is best effort: a failure never stops
' the user's own work. Times come from each computer's clock.
Public Const CONNECTED_MINUTES As Long = 3
Private Const EDITING_MINUTES As Long = 3
Private Const TOUCH_SECONDS As Long = 30
Private lastTouch As Single

' This session's identity; TempVars keep it through a VBA reset.
Public Function PresenceSession() As String
    On Error Resume Next
    PresenceSession = CStr(TempVars("SwiftSessionGuid").Value)
    If Len(PresenceSession) = 0 Then
        Err.Clear
        TempVars.Add "SwiftSessionGuid", NewIdentity()
        PresenceSession = CStr(TempVars("SwiftSessionGuid").Value)
    End If
End Function

' Record that this session is here. At most every 30 seconds unless forced.
Public Sub PresenceTouch(Optional ByVal force As Boolean = False)
    Dim r As DAO.Recordset, session As String, attempt As Long
    On Error GoTo Done
    If Not force And lastTouch > 0 And Abs(Timer - lastTouch) < TOUCH_SECONDS Then Exit Sub
    If Not HasActiveContributor() Then Exit Sub
    session = PresenceSession()
    Set r = CurrentDb.OpenRecordset("SELECT * FROM SessionPresence WHERE SessionGuid=" & QuoteSQL(session), dbOpenDynaset)
    If r.EOF Then
        r.AddNew: r!SessionGuid = session: r!StartedAt = Now()
    Else
        r.Edit
    End If
    r!ProfileID = ActiveContributorID(): r!ComputerName = Left$(Environ$("COMPUTERNAME"), 64): r!LastSeen = Now()
    r.Update: r.Close: Set r = Nothing
    lastTouch = Timer
    'Sessions that ended without saying so.
    CurrentDb.Execute "DELETE FROM SessionPresence WHERE LastSeen < DateAdd('h',-2,Now())", dbFailOnError
Done:
    'Another session writing the table at that moment: retry the statement a few times.
    If SwiftLockConflict(Err.Number) And attempt < 3 Then
        attempt = attempt + 1
        SwiftLockWait attempt
        Resume
    End If
    On Error Resume Next
    If Not r Is Nothing Then r.Close
End Sub

' The report this session has open in its editor (0 when none).
Public Sub PresenceEditing(ByVal reportID As Long)
    On Error GoTo Done
    PresenceTouch True
    If reportID > 0 Then
        PresenceWrite "UPDATE SessionPresence SET ReportID=" & reportID & ", EditingSince=Now() WHERE SessionGuid=" & QuoteSQL(PresenceSession())
    Else
        PresenceWrite "UPDATE SessionPresence SET ReportID=Null, EditingSince=Null WHERE SessionGuid=" & QuoteSQL(PresenceSession())
    End If
Done:
End Sub

' The keeper form's timer and unload events (form events call functions).
Public Function PresenceKeep() As Boolean
    PresenceTouch True: PresenceKeep = True
End Function
Public Function PresenceLeave() As Boolean
    Call PresenceEnd: PresenceLeave = True
End Function
' Start the hidden keeper for this session, once a contributor is chosen.
Public Sub PresenceStartKeeper()
    On Error Resume Next
    If Not CurrentProject.AllForms("frmSessionKeeper").IsLoaded Then DoCmd.OpenForm "frmSessionKeeper", , , , , acHidden
End Sub

' This session is leaving.
Public Sub PresenceEnd()
    On Error Resume Next
    PresenceWrite "DELETE FROM SessionPresence WHERE SessionGuid=" & QuoteSQL(PresenceSession())
End Sub

' Presence writes are plain and best effort; other sessions read them fresh.
Private Sub PresenceWrite(ByVal sql As String)
    On Error Resume Next
    CurrentDb.Execute sql, dbFailOnError
End Sub

Private Function OtherSessions(ByVal extra As String, ByVal minutes As Long, Optional ByRef count As Long) As String
    Dim r As DAO.Recordset, names As String, entry As String
    count = 0
    On Error GoTo Done
    'Other sessions' changes are cached for a few seconds; read them fresh.
    DBEngine.Idle dbRefreshCache
    Set r = CurrentDb.OpenRecordset("SELECT s.ComputerName, p.DisplayName FROM SessionPresence AS s LEFT JOIN UserProfile AS p ON s.ProfileID=p.UserProfileID WHERE s.SessionGuid<>" & QuoteSQL(PresenceSession()) & " AND s.LastSeen>=DateAdd('n'," & -minutes & ",Now())" & extra & " ORDER BY s.LastSeen DESC", dbOpenSnapshot)
    Do While Not r.EOF
        entry = Nz(r!DisplayName, "Someone") & " (" & Nz(r!ComputerName, "another computer") & ")"
        count = count + 1
        If InStr(", " & names & ",", ", " & entry & ",") = 0 Then names = names & IIf(Len(names) > 0, ", ", "") & entry
        r.MoveNext
    Loop
    r.Close
    OtherSessions = names
Done:
End Function

' Other sessions connected now, and their names with computers.
Public Function PresenceOthersNow(ByRef names As String) As Long
    Dim count As Long
    names = OtherSessions("", CONNECTED_MINUTES, count)
    PresenceOthersNow = count
End Function

' The Overview line: who else has this workspace open now.
Public Function PresenceOthersText() As String
    Dim names As String
    If PresenceOthersNow(names) > 0 Then PresenceOthersText = "Also open: " & names & "."
End Function

' Someone else working under this session's contributor name.
Public Function PresenceSameNameText() As String
    Dim names As String
    If Not HasActiveContributor() Then Exit Function
    names = OtherSessions(" AND s.ProfileID=" & ActiveContributorID(), CONNECTED_MINUTES)
    If Len(names) > 0 Then PresenceSameNameText = names & " is also working under your name. Use your own name so your work and theirs stay apart."
End Function

' The editor notice: who else has this report open right now.
Public Function PresenceEditingText(ByVal reportID As Long) As String
    Dim names As String
    If reportID <= 0 Then Exit Function
    names = OtherSessions(" AND s.ReportID=" & reportID, EDITING_MINUTES)
    If Len(names) > 0 Then PresenceEditingText = names & " also has this report open. If they save first, your save is refused and your entries are kept."
End Function
