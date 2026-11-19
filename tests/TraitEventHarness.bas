Attribute VB_Name = "TraitEventHarness"
Option Explicit
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Public Function TestTraitEvents(Optional ByVal injectLock As Boolean = False) As String
 Dim groupID As Long, id As Long, rows As Access.Form, revision As Long, kind As Variant, grade As String, rate As String
 Dim locks As DAO.Workspace, lockDB As DAO.Database, locked As DAO.Recordset, status As String
 On Error GoTo Failed
 ConfigureWorkspace "Event test", "Tester"
 For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
  Select Case CStr(kind)
   Case "EVAL": grade = "E5": rate = "IT2"
   Case "CHIEFEVAL": grade = "E7": rate = "ITC"
   Case Else: grade = "O3": rate = "LT"
  End Select
  groupID = CreateSummaryGroup("New " & kind, CStr(kind), grade)
  id = CreateGroupDraft(groupID, GroupStamp(groupID), UXDefaultsStamp(groupID), 0, 0, "Synthetic event", "EVENT, SYNTHETIC", rate, "", Null, Null)
  DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = groupID
  UXWorkbenchMode "Members": UXOpenComparison
  Set rows = Forms!frmUXComparison!uxTraitGrid.Form
  TraceEvent "start " & kind
  revision = CurrentRevision(id)
  rows!Trait33.SetFocus: rows!Trait33.Text = "4.0": rows!Trait34.SetFocus
  If CurrentRevision(id) <> revision Then Err.Raise vbObjectError + 7907, , "Report transaction ran inside control event"
  TraceEvent "before drain"
  If Not AwaitTraitSave(id, "Trait33", "4.0") Then Err.Raise vbObjectError + 7908, , Forms!frmUXComparison!lblStatus.Caption
  TraceEvent "after drain"
  If CurrentRevision(id) = revision Then Err.Raise vbObjectError + 7901, , Forms!frmUXComparison!lblStatus.Caption
  If DLookup("Trait33", "Report", "ReportID=" & id) <> "4.0" Then Err.Raise vbObjectError + 7902, , "Trait reverted"
  rows!Trait34.Text = "3.0": rows!Trait35.SetFocus
  TraceEvent "before drain"
  If Not UXGridDrain() Then Err.Raise vbObjectError + 7908, , Forms!frmUXComparison!lblStatus.Caption
  If DLookup("Trait34", "Report", "ReportID=" & id) <> "3.0" Then Err.Raise vbObjectError + 7902, , "Second trait reverted"
  rows!Trait35.Text = "1.0": rows!Trait36.SetFocus
  TraceEvent "before close"
  UXCloseComparison
  TraceEvent "after close"
  If DLookup("Trait35", "Report", "ReportID=" & id) <> "1.0" Then Err.Raise vbObjectError + 7909, , "Close lost queued edit"
 Next
 If Not injectLock Then TestTraitEvents = "PASS: repeated native control events for all three report types": Exit Function
 Call UXOpenComparison: Set rows = Forms!frmUXComparison!uxTraitGrid.Form
 Set locks = DBEngine.CreateWorkspace("TestExternalReportLock", "admin", "", dbUseJet)
 Set lockDB = locks.OpenDatabase(CurrentProject.FullName)
 Set locked = lockDB.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & id, dbOpenDynaset)
 locked.LockEdits = True: locked.Edit: locked!Title = "Uncommitted external edit"
 revision = CurrentRevision(id)
 rows!Trait33 = "2.0"
 If UXGridEdit("Trait33") Then Err.Raise vbObjectError + 7903, , "External lock save unexpectedly succeeded"
 status = Forms!frmUXComparison!lblStatus.Caption
 If CurrentRevision(id) <> revision Then Err.Raise vbObjectError + 7904, , "Lock failure created a revision"
 locked.CancelUpdate: locked.Close: lockDB.Close: locks.Close
 rows!Trait33 = "2.0"
 If Not UXGridEdit("Trait33") Then Err.Raise vbObjectError + 7905, , "Retry failed: " & Forms!frmUXComparison!lblStatus.Caption
 If DLookup("Trait33", "Report", "ReportID=" & id) <> "2.0" Then Err.Raise vbObjectError + 7906, , "Retry not persisted"
 TestTraitEvents = "PASS all three native cell events, external lock rollback and retry; caught: " & status: Exit Function
Failed:
 TestTraitEvents = "FAIL: " & Err.Number & " " & Err.Description
 On Error Resume Next
 locked.CancelUpdate: locked.Close: lockDB.Close: locks.Close
End Function

Private Sub TraceEvent(ByVal message As String)
 Dim n As Integer
 n = FreeFile
 Open CurrentProject.FullName & ".trace" For Append As #n
 Print #n, message
 Close #n
End Sub

Private Function AwaitTraitSave(ByVal id As Long, ByVal key As String, ByVal mark As String) As Boolean
 Dim started As Single
 started = Timer
 Do
  DoEvents
  If Nz(DLookup(key, "Report", "ReportID=" & id), "") = mark Then AwaitTraitSave = True: Exit Function
  Sleep 10
 Loop While Timer - started < 3
End Function
