Attribute VB_Name = "modSessionPresence"
Option Compare Database
Option Explicit
' Who else has this workspace open (backlog S01: two people on one data file, not
' supported but made safer). Each session keeps one SessionPresence row: its contributor,
' computer, when it started, when it was last seen, and the report open in its editor.
' Rows not seen recently are ignored, and old ones are cleared. Everything here is best
' effort: a failure never stops the user's own work.
Private Const PRESENT_MINUTES As Long = 10
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
    Dim r As DAO.Recordset, session As String
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

Private Function OtherSessions(ByVal extra As String, ByVal minutes As Long) As String
    Dim r As DAO.Recordset, names As String, entry As String
    On Error GoTo Done
    'Other sessions' changes are cached for a few seconds; read them fresh.
    DBEngine.Idle dbRefreshCache
    Set r = CurrentDb.OpenRecordset("SELECT s.ComputerName, p.DisplayName FROM SessionPresence AS s LEFT JOIN UserProfile AS p ON s.ProfileID=p.UserProfileID WHERE s.SessionGuid<>" & QuoteSQL(PresenceSession()) & " AND s.LastSeen>=DateAdd('n'," & -minutes & ",Now())" & extra & " ORDER BY s.LastSeen DESC", dbOpenSnapshot)
    Do While Not r.EOF
        entry = Nz(r!DisplayName, "Someone") & " (" & Nz(r!ComputerName, "another computer") & ")"
        If InStr(", " & names & ",", ", " & entry & ",") = 0 Then names = names & IIf(Len(names) > 0, ", ", "") & entry
        r.MoveNext
    Loop
    r.Close
    OtherSessions = names
Done:
End Function

' The Overview line: who else has used this workspace in the last few minutes.
Public Function PresenceOthersText() As String
    Dim names As String
    names = OtherSessions("", PRESENT_MINUTES)
    If Len(names) > 0 Then PresenceOthersText = "Also used in the last " & PRESENT_MINUTES & " minutes: " & names & ". Work on different reports; a save that would overwrite someone else's is refused."
End Function

' The editor notice: who else has this report open right now.
Public Function PresenceEditingText(ByVal reportID As Long) As String
    Dim names As String
    If reportID <= 0 Then Exit Function
    names = OtherSessions(" AND s.ReportID=" & reportID, EDITING_MINUTES)
    If Len(names) > 0 Then PresenceEditingText = names & " also has this report open. If they save first, your save is refused and your entries are kept."
End Function
