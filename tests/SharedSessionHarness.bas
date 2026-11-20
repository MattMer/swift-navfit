Attribute VB_Name = "SharedSessionHarness"
Option Compare Database
Option Explicit
' Several people on one data file (scripts/cases/test-shared-sessions.ps1). The
' coordinator calls SharedSetup once; then each worker process runs SharedWorker in its
' own Access session against the same data file, phase by phase at shared start times.
' Workers report outcomes; the coordinator checks the data file itself.

Private Function WaitUntil(ByVal t0 As Double, ByVal seconds As Long) As Boolean
    Dim target As Date
    target = CDate(t0) + seconds / 86400#
    Do While Now() < target
        DoEvents
        SwiftPause 50
    Loop
    WaitUntil = True
End Function

Private Function Changes(ByVal key As String, ByVal value As Variant) As Object
    Set Changes = CreateObject("Scripting.Dictionary")
    Changes.Add key, value
End Function

' One shared data file with eight contributors, a report for each, one report everyone
' reaches for, one member and one reporting senior. Returns the ids, "|"-separated:
' the report everyone edits, member, senior, then profile;report for workers 1 to 8.
Public Function SharedSetup(ByVal dataPath As String) As String
    Dim ids As String, i As Long, profile As Long, report As Long, member As Long, senior As Long, everyone As Long, values As Object
    On Error GoTo Failed
    If Not CreateDataFile(dataPath) Then Err.Raise vbObjectError + 9700, , "Data file not created"
    LinkDataFile dataPath
    ConfigureWorkspace "SYNTHETIC SHARED COMMAND", "Coordinator"
    everyone = CommitDraft(0, "EVAL", "Report everyone edits", "Shared Report", #1/1/2026#, #9/30/2026#, "", "Test")
    member = SaveDirectoryMember(0, 0, "Shared Member", "SHARED, MEMBER", "IT2", "000000071")
    Set values = RSSeniorValues(0): values("ReportingSeniorSSN") = "000123471": values("ReportingSeniorName") = "SHARED, RS"
    senior = RSSaveSenior(0, "Shared senior", values)
    ids = everyone & "|" & member & "|" & senior
    For i = 1 To 8
        profile = AddContributor("Worker " & i)
        report = CommitDraft(0, "EVAL", "Own report " & i, "Worker Member " & i, #1/1/2026#, #9/30/2026#, "", "Test")
        ids = ids & "|" & profile & ";" & report
    Next
    SharedSetup = "PASS:" & ids
    Exit Function
Failed:
    SharedSetup = "FAIL: " & Err.Number & " " & Err.Description
End Function

Private Function Outcome() As String
    Select Case Err.Number
        Case vbObjectError + 900, vbObjectError + 4004, vbObjectError + 4806, vbObjectError + 2011, vbObjectError + 2012
            Outcome = "stale"
        Case Else
            If InStr(Err.Description, "changed") > 0 Or InStr(Err.Description, "newer revision") > 0 Then Outcome = "stale" Else Outcome = "error:" & Err.Number & " " & Err.Description
    End Select
End Function

' One worker's run. Phases start at t0 + 0 (startup), 25 (four saves at a working pace),
' 110 (all save at the same second), 140 (contest, from the version all read at 135)
' and 160 (who is connected).
Public Function SharedWorker(ByVal worker As Long, ByVal idList As String, ByVal t0 As Double) As String
    Dim ids As Variant, mine As Variant, out As String, i As Long, base As Long, ver As Long, names As String, count As Long, own As Long, everyone As Long, started As Single, times As String
    ids = Split(idList, "|"): mine = Split(ids(2 + worker), ";")
    everyone = CLng(ids(0)): own = CLng(mine(1))
    'Phase 1: everyone starts at once.
    WaitUntil t0, 0
    On Error Resume Next
    out = "ready=" & DataLinkReady() & IIf(Err.Number <> 0, " error:" & Err.Number & " " & Err.Description, "")
    Err.Clear
    BindActiveContributor CLng(mine(0))
    If Err.Number <> 0 Then out = out & " bind-error:" & Err.Description
    out = out & vbCrLf & "hold=" & Left$(SwiftRowLockHeld(), 4)
    Err.Clear
    'Phase 2: each saves its own report four times at a working pace while the others
    'do the same, then all eight save at the same second.
    WaitUntil t0, 25
    Randomize
    For i = 1 To 4
        SwiftPause 8000 + CLng(Rnd() * 7000)
        SwiftFreshRead
        started = Timer
        CommitSavedFields own, CurrentRevision(own), Changes("Title", "Own report " & worker & " save " & i), "Shared test"
        times = times & IIf(Len(times) > 0, ",", "") & Format$(Timer - started, "0.0")
        If Err.Number <> 0 Then out = out & vbCrLf & "own-save-" & i & "=" & Outcome(): Err.Clear
        'Worker 2 makes a backup while the others save: verified, or refused with nothing left.
        If worker = 2 And i = 2 Then
            started = Timer
            WriteWorkspaceBackup CurrentProject.Path & "\backup-during-saves.accdb"
            out = out & vbCrLf & "backup=" & IIf(Err.Number = 0, "done", IIf(Err.Number = vbObjectError + 4831, "refused", "error:" & Err.Number & " " & Err.Description)) & " in " & Format$(Timer - started, "0.0") & " s"
            Err.Clear
        End If
    Next
    If Now() > CDate(t0) + 109 / 86400# Then out = out & vbCrLf & "error:late for the same-second save"
    SwiftFreshRead
    i = CurrentRevision(own)
    WaitUntil t0, 110
    started = Timer
    CommitSavedFields own, i, Changes("Title", "Own report " & worker & " burst"), "Shared test"
    times = times & "," & Format$(Timer - started, "0.0")
    If Err.Number <> 0 Then out = out & vbCrLf & "own-save-burst=" & Outcome(): Err.Clear
    out = out & vbCrLf & "seconds=" & times
    out = out & vbCrLf & "own=done"
    'Phase 3: everyone starts from the same saved version of a shared record.
    If Now() > CDate(t0) + 134 / 86400# Then out = out & vbCrLf & "error:late for the shared contest"
    WaitUntil t0, 135
    SwiftFreshRead
    Select Case worker
        Case 1 To 4
            base = CurrentRevision(everyone)
            WaitUntil t0, 140
            CommitSavedFields everyone, base, Changes("DraftNotes", "Saved by worker " & worker), "Shared test"
        Case 5, 6
            ver = CLng(DLookup("VersionCounter", "CommandMember", "MemberID=" & CLng(ids(1))))
            WaitUntil t0, 140
            SaveDirectoryMember CLng(ids(1)), ver, "Shared Member " & worker, "SHARED, MEMBER", "IT2", "000000071"
        Case 7, 8
            ver = CLng(DLookup("VersionCounter", "ReportingSeniorRecord", "SeniorID=" & CLng(ids(2))))
            WaitUntil t0, 140
            RSSaveSenior CLng(ids(2)), "Shared senior " & worker, RSSeniorValues(CLng(ids(2))), ver
    End Select
    out = out & vbCrLf & "contest=" & IIf(Err.Number = 0, "won", Outcome())
    Err.Clear
    'Phase 4: everyone sees the other seven; then worker 8 leaves and worker 1 sees six.
    WaitUntil t0, 160
    PresenceTouch True
    WaitUntil t0, 167
    count = PresenceOthersNow(names)
    out = out & vbCrLf & "others=" & count
    If worker = 8 Then WaitUntil t0, 171: DoCmd.Close acForm, "frmSessionKeeper", acSaveNo
    If worker = 1 Then WaitUntil t0, 177: out = out & vbCrLf & "after-leave=" & PresenceOthersNow(names)
    'Compacting while others are connected is refused, naming them.
    If worker = 1 Then
        Err.Clear
        CompactDataFile
        out = out & vbCrLf & "compact=" & IIf(Err.Number = 0, "done", IIf(InStr(Err.Description, "Open now: Worker") > 0, "refused", "error:" & Err.Number & " " & Err.Description))
        Err.Clear
    End If
    'The others stay open until worker 1 has looked.
    If worker <> 8 Then WaitUntil t0, 180
    If Err.Number <> 0 Then out = out & vbCrLf & "error:" & Err.Number & " " & Err.Description
    SharedWorker = out
End Function

' Once everyone else has closed: compacts the data file with Overview open, as a
' coordinator would. Returns "PASS:<result>" or the error.
Public Function SharedCompact(ByVal profile As Long) As String
    On Error GoTo Failed
    If Not DataLinkReady() Then SharedCompact = "FAIL: data file not ready": Exit Function
    BindActiveContributor profile
    DoCmd.OpenForm "frmHome"
    SharedCompact = "PASS:" & CompactDataFile()
    If CurrentProject.AllForms("frmHome").IsLoaded Then DoCmd.Close acForm, "frmHome", acSaveNo
    Exit Function
Failed:
    SharedCompact = "FAIL: " & Err.Number & " " & Err.Description
End Function

' How long one save of a report takes with nobody else working: n saves, the average
' and the slowest, in milliseconds.
Public Function SharedTimeSave(ByVal reportID As Long, ByVal n As Long) As String
    Dim i As Long, started As Single, took As Single, total As Single, slowest As Single
    On Error GoTo Failed
    For i = 1 To n
        started = Timer
        CommitSavedFields reportID, CurrentRevision(reportID), Changes("Title", "Timing save " & i), "Timing"
        took = (Timer - started) * 1000: total = total + took
        If took > slowest Then slowest = took
    Next
    SharedTimeSave = "average " & Format$(total / n, "0") & " ms, slowest " & Format$(slowest, "0") & " ms over " & n & " saves"
    Exit Function
Failed:
    SharedTimeSave = "error:" & Err.Number & " " & Err.Description
End Function
