Attribute VB_Name = "ContributorStartupHarness"
Option Compare Binary
Option Explicit
Private checks As Long
Private Sub Check(ByVal condition As Boolean, ByVal message As String)
 If Not condition Then Err.Raise vbObjectError + 8910, , message
 checks = checks + 1
End Sub
Private Sub RequiresSelection(ByVal action As String, Optional ByVal report As Long = 0)
 Dim failure As Long
 On Error Resume Next
 'Call VBA services directly: nested Access Eval callbacks bypass the caller's trap.
 Select Case action
  Case "Save": Call CommitDraft(0, "EVAL", "Blocked", "Member", Null, Null, "")
  Case "Operation": Call BeginOperation("Blocked", "Test")
  Case "Message": Call AddReportMessage(report, "Blocked message")
  Case "Assign": Call AssignReview(report, 0, 1, Null, "Blocked review")
  Case "Decision": Call FinishReview(0, "Accepted", "Blocked decision")
  Case "Add": Call AddContributor("Blocked extra person")
 End Select
 failure = Err.Number
 Err.Clear
 On Error GoTo 0
 Check failure = vbObjectError + 8900, "Unselected action did not fail closed: " & action & " (" & failure & ")"
End Sub
Private Sub UnselectedChecks(ByVal report As Long)
 Check Not HasActiveContributor(), "Stored contributor silently accepted"
 Check ActiveContributorCaption() = "Choose a contributor", "Retained identity displayed as confirmed"
 RequiresSelection "Save"
 RequiresSelection "Operation"
 RequiresSelection "Message", report
 RequiresSelection "Assign", report
 RequiresSelection "Decision"
 RequiresSelection "Add"
 Check DCount("*", "Report", "Title='Blocked'") = 0, "Blocked report committed"
 Check DCount("*", "OperationJournal", "OperationKind='Blocked'") = 0, "Blocked operation committed"
End Sub
Public Function TestContributorStartup() As String
 Dim first As Long, second As Long, added As Long, report As Long, message As Long, task As Long, oldRevision As Long, stage As String
 Dim db As DAO.Database, oldRule As String
 On Error GoTo Failed
 stage = "blank startup"
 Check DCount("*", "Workspace") = 0, "Test requires blank fixture"
 Check CurrentProject.AllForms("frmSetup").IsLoaded, "Blank startup omitted setup"
 Check Not HasActiveContributor(), "Blank startup selected a person"
 ConfigureWorkspace "SYNTHETIC CONTRIBUTOR TEST", "First contributor"
 first = ActiveContributorID()
 Check HasActiveContributor(), "Explicit setup did not select creator"
 second = AddContributor("Second contributor")
 report = CommitDraft(0, "EVAL", "Contributor test", "Synthetic member", Null, Null, "")
 oldRevision = CurrentRevision(report)
 Check DLookup("AuthorProfileID", "ReportRevision", "RevisionID=" & oldRevision) = first, "First author incorrect"
 stage = "initialized startup"
 Check SwiftAppStartup(), "Startup failed"
 Check CurrentProject.AllForms("frmContributorStart").IsLoaded, "Chooser missing"
 Check Not CurrentProject.AllForms("frmHome").IsLoaded, "Home opened before selection"
 Check IsNull(Forms!frmContributorStart!cboContributor), "Chooser defaulted retained name"
 Check Forms!frmContributorStart!cboContributor.ListCount = 2, "Contributor list incomplete"
 UnselectedChecks report
 Check Not AcceptContributorStartup(), "Empty selection accepted"
 Check DLookup("ActiveProfileID", "Workspace") = first, "Failed choice changed retained profile"
 Forms!frmContributorStart!txtNewContributor = "   "
 Check Not AcceptContributorStartup(True), "Whitespace name accepted"
 Forms!frmContributorStart!txtNewContributor = String$(121, "x")
 Check Not AcceptContributorStartup(True), "Oversized name accepted"
 Forms!frmContributorStart!txtNewContributor = "First contributor"
 Check Not AcceptContributorStartup(True), "Duplicate name accepted"
 Check DCount("*", "UserProfile") = 2, "Invalid add created a profile"
 Forms!frmContributorStart!cboContributor = second
 Check AcceptContributorStartup(), "Existing choice rejected"
 Check ActiveContributorID() = second, "Existing choice not bound"
 Check CurrentProject.AllForms("frmHome").IsLoaded, "Home missing after selection"
 Check InStr(Forms!frmHome!uxContext.Caption, "Second contributor") > 0, "Selected name not displayed"
 stage = "attribution"
 report = CommitDraft(report, "EVAL", "Contributor test", "Synthetic member", Null, Null, "Second author's edit")
 Check DLookup("AuthorProfileID", "ReportRevision", "RevisionID=" & CurrentRevision(report)) = second, "Save used old actor"
 Check DLookup("AuthorProfileID", "ReportRevision", "RevisionID=" & oldRevision) = first, "Historical author changed"
 message = AddReportMessage(report, "Second contributor's note")
 Check DLookup("AuthorProfileID", "ReportMessage", "MessageID=" & message) = second, "Note used old actor"
 task = AssignReview(report, CurrentRevision(report), first, Null, "Synthetic assignment")
 Check DLookup("RequestedBy", "ReviewTask", "TaskID=" & task) = second, "Review used old actor"
 Check SwitchContributor(first), "Reviewer switch failed"
 Check FinishReview(task, "Accepted", "Synthetic review"), "Review decision failed"
 Check DLookup("AuthorProfileID", "ReviewTaskEvent", "TaskID=" & task & " AND EventKind='Accepted'") = first, "Review decision actor wrong"
 stage = "new contributor restart"
 Check SwiftAppStartup(), "Repeat startup failed"
 stage = "new contributor entry"
 Set db = CurrentDb
 oldRule = Nz(db.TableDefs("Workspace").ValidationRule, "")
 'Inject an update rejection only in this disposable fixture.
 db.TableDefs("Workspace").ValidationRule = "[ActiveProfileID]=" & first
 Forms!frmContributorStart!txtNewContributor = "Rollback contributor"
 Check Not AcceptContributorStartup(True), "Rejected selection committed new profile"
 db.TableDefs("Workspace").ValidationRule = oldRule
 Check DCount("*", "UserProfile") = 2, "New-profile failure did not roll back"
 Check Not HasActiveContributor(), "New-profile rollback confirmed session"
 Check DLookup("ActiveProfileID", "Workspace") = first, "New-profile rollback changed retained selection"
 Forms!frmContributorStart!txtNewContributor = "  O'Brien contributor  "
 stage = "new contributor acceptance"
 If Not AcceptContributorStartup(True) Then Err.Raise vbObjectError + 8910, , Forms!frmContributorStart!lblStatus.Caption
 Check HasActiveContributor(), "New contributor not confirmed"
 stage = "new contributor identity"
 added = ActiveContributorID()
 Check added <> first And added <> second, "New profile not selected"
 Check ActiveContributorCaption() = "O'Brien contributor", "New name not trimmed/preserved"
 Check DCount("*", "UserProfile") = 3, "New contributor count wrong"
 stage = "another session's choice"
 'Another session choosing a different contributor leaves this session's own in place.
 CurrentDb.Execute "UPDATE Workspace SET ActiveProfileID=" & first & " WHERE WorkspaceID=1", dbFailOnError
 Check HasActiveContributor() And ActiveContributorID() = added, "Another session's choice replaced this session's contributor"
 Call SwitchContributor(added)
 TempVars.Remove "SwiftContributorPath"
 Check Not HasActiveContributor(), "Incomplete session accepted"
 Check ContributorSessionProblem() = "Choose your contributor for this session before making changes.", "Incomplete session blamed another session"
 Call SwitchContributor(added)
 Check SwiftAppStartup(), "Final startup failed"
 Check IsNull(Forms!frmContributorStart!cboContributor), "Repeat startup remembered selection"
 UnselectedChecks report
 TestContributorStartup = "PASS: " & checks & " contributor startup, validation, attribution and fail-closed assertions"
 Exit Function
Failed:
 TestContributorStartup = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description & " / " & ContributorStartupFailure()
End Function

'Startup warnings: classify synthetic paths, a real read-only file, the session
'count of this exclusive fixture and the clock-safe checkpoint schedule.
Public Function TestWorkspaceSafety() As String
 Dim stage As String, folder As String, path As String, computers As String, t As Date, text As String
 On Error GoTo Failed
 checks = 0
 stage = "location"
 Check InStr(WorkspaceLocationWarning(Environ$("TEMP") & "\Temp1_swift.zip\Swift.accdb"), "temporary folder") > 0, "Zip extraction folder not flagged"
 Check InStr(WorkspaceLocationWarning("C:\Users\Sailor\AppData\Local\Microsoft\Windows\INetCache\Content.Outlook\AB12CD\Swift.accdb"), "temporary folder") > 0, "E-mail attachment folder not flagged"
 Check InStr(WorkspaceLocationWarning("C:\Users\Sailor\OneDrive - US Navy\Reports\Swift.accdb"), "sync service") > 0, "Commercial OneDrive folder not flagged"
 Check InStr(WorkspaceLocationWarning("C:\Users\Sailor\Dropbox\Swift.accdb"), "sync service") > 0, "Dropbox folder not flagged"
 If Len(Environ$("OneDrive")) > 0 Then Check InStr(WorkspaceLocationWarning(Environ$("OneDrive") & "\Swift.accdb"), "sync service") > 0, "Configured OneDrive root not flagged"
 Check WorkspaceLocationWarning("D:\Command\Evaluations\Swift.accdb") = "", "Ordinary folder flagged"
 Check WorkspaceLocationWarning("\fileserver\share\Evaluations\Swift.accdb") = "", "Network share flagged as unsafe location"
 Check WorkspaceLocationWarning("D:\Templates\TempLists\Swift.accdb") = "", "Folder name containing Temp flagged"
 stage = "read-only file"
 folder = CurrentProject.Path & "\safety-" & Replace(Replace(NewIdentity(), "{", ""), "}", "")
 MkDir folder
 path = folder & "\readonly.accdb"
 Open path For Output As #1: Print #1, "synthetic": Close #1
 Check WorkspaceFileWarning(path) = "", "Writable ordinary file flagged"
 SetAttr path, vbReadOnly
 Check InStr(WorkspaceFileWarning(path), "read-only") > 0, "Read-only file not flagged"
 SetAttr path, vbNormal: Kill path: RmDir folder
 path = "": folder = ""
 stage = "sessions"
 Check OtherWorkspaceSessions(computers) = 0 And computers = "", "Exclusive fixture reported another session: " & computers
 Check ConcurrentSessionWarning(0, "") = "", "No other session produced a warning"
 Check InStr(ConcurrentSessionWarning(1, "PC-2"), "already open in another session (PC-2)") > 0, "One other session not named"
 'Copies a sync service made beside the data file are named; Swift's own copies are not.
 Dim syncFolder As String, own As String
 syncFolder = Environ$("TEMP") & "\SwiftSync" & Format$(Now(), "hhnnss") & "\": MkDir syncFolder
 own = syncFolder & "Swift NAVFIT data.accdb"
 Open own For Output As #1: Close #1
 Check SyncConflictWarning(own) = "", "A lone data file was flagged"
 Open syncFolder & "Swift NAVFIT data before schema 49 20261113-120000.accdb" For Output As #1: Close #1
 Open syncFolder & "Swift NAVFIT data 2.accdb" For Output As #1: Close #1
 Check SyncConflictWarning(own) = "", "Swift's own copies were flagged"
 Open syncFolder & "Swift NAVFIT data-PC-2.accdb" For Output As #1: Close #1
 Open syncFolder & "Swift NAVFIT data (1).accdb" For Output As #1: Close #1
 Check InStr(SyncConflictWarning(own), "Swift NAVFIT data-PC-2.accdb") > 0 And InStr(SyncConflictWarning(own), "Swift NAVFIT data (1).accdb") > 0, "Sync conflict copies not named: " & SyncConflictWarning(own)
 Kill syncFolder & "*.accdb": RmDir syncFolder
 Check InStr(ConcurrentSessionWarning(2, "PC-2, this computer"), "2 other sessions") > 0, "Session count not reported"
 stage = "notice"
 text = WorkspaceSafetyNotice()
 Check InStr(text, "could not check") = 0, "Safety probe failed: " & text
 Check InStr(text, "already open") = 0, "Exclusive fixture warned about another session"
 stage = "checkpoint schedule"
 t = #11/6/2026 10:00:00 AM#
 Check CheckpointSchedule(0, t) = DateAdd("s", 30, t), "First checkpoint not scheduled"
 Check CheckpointSchedule(DateAdd("s", 10, t), t) = DateAdd("s", 10, t), "Pending checkpoint moved"
 Check CheckpointSchedule(DateAdd("s", 30, t), t) = DateAdd("s", 30, t), "Full-interval checkpoint moved"
 Check CheckpointSchedule(DateAdd("h", 1, t), t) = DateAdd("s", 30, t), "Clock set back an hour stalled checkpoints"
 Check CheckpointSchedule(DateAdd("h", -1, t), t) = DateAdd("h", -1, t), "Clock moved forward did not leave the checkpoint due"
 TestWorkspaceSafety = "PASS: " & checks & " workspace location, read-only, session and checkpoint-schedule assertions"
 Exit Function
Failed:
 TestWorkspaceSafety = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
 'On Error inside an active handler is ignored; leave the handler before cleanup.
 Resume Cleanup
Cleanup:
 On Error Resume Next
 If Len(path) > 0 Then
  If Len(Dir$(path)) > 0 Then SetAttr path, vbNormal: Kill path
 End If
 If Len(folder) > 0 Then RmDir folder
End Function
