Attribute VB_Name = "EfficiencyHarness"
Option Explicit
Private checks As Long, timerID As Long, paintBefore As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8060, , message
 checks = checks + 1
End Sub
Public Function TestEfficiency() As String
 Dim id As Long, groupID As Long, f As Form, count As Long, average As String, i As Long
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC EFFICIENCY", "Tester"
 id = CommitDraft(0, "EVAL", "Synthetic", "Synthetic", Null, Null, "", "Test", Null, , Array("4.0", "4.0", "4.0", "4.0", "4.0", "4.0", "4.0"), Array("Observed", "Promotable"))
 groupID = CreateSummaryGroup("Direct navigation", "EVAL", "E5"): AddGroupMember groupID, id, "E5"
 DoCmd.OpenForm "frmEditor", , , , , , CStr(id): Set f = Forms!frmEditor
 f.TimerInterval = 0: UXRefreshEditor: count = UXEditorGroupRefreshCount(): average = f!uxSGA.Caption
 f!txtOfficialComments = "Pending narrative": Call UXEditorChanged(): UXRefreshEditor "OfficialComments"
 Check UXEditorGroupRefreshCount() = count, "Narrative edit refreshed group data"
 Check f!uxSGA.Caption = average, "Narrative changed saved group average"
 f!cboTrait33 = "3.0": Call UXEditorChanged(): UXRefreshEditor "Trait33"
 Check UXEditorGroupRefreshCount() = count, "Trait edit refreshed saved group data"
 Check InStr(f!uxITA.Caption, "3.86") > 0, "Pending ITA not refreshed"
 Check f!uxSGA.Caption = average, "Pending trait changed saved SGA"
 UXRefreshEditor: Check UXEditorGroupRefreshCount() = count, "Unchanged full refresh repeated group query"
 count = UXLivePaintCount()
 For i = 1 To 20: Call UXLiveTimer(): Next
 Check UXLivePaintCount() = count, "Idle timer scanned editor fields"
 f!txtMemberSSN = "INVALID": Call UXLiveTyped()
 Check UXLivePaintCount() = count, "Typing validated synchronously"
 Call UXLiveTimer(): Check UXLivePaintCount() = count + 1, "Typing did not trigger next timer validation"
 f!txtMemberSSN = Null: Call UXEditorChanged(): UXRefreshEditor "MemberSSN"
 count = UXEditorGroupRefreshCount(): Check UXSaveEditor(), "Save failed"
 Check UXEditorGroupRefreshCount() = count + 1, "Context refresh unavailable"
 Check InStr(f!uxSGA.Caption, "3.86") > 0, "Saved group average stale after commit"
 f.OnUnload = "": DoCmd.Close acForm, "frmEditor", acSaveNo
 If CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then DoCmd.Close acForm, "frmMajorWorkbench", acSaveNo
 DoCmd.OpenForm "frmUXGroups": Forms!frmUXGroups!lstGroups = groupID
 Check UXOpenGroup(), "Direct group open failed"
 Check Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded, "Members opened legacy context host"
 Check Not UXHasControl(Forms!frmUXComparison, "lstTraits"), "Duplicate comparison list retained"
 Check UXGroupNavigate("frmUXComparison", 3), "Comparison tab failed"
 Check DCount("*", "GroupSnapshot") = 0, "Tab switch created history"
 Check UXGroupNavigate("frmUXComparison", 1), "Defaults navigation failed"
 Check CurrentProject.AllForms("frmUXBulk").IsLoaded, "Defaults not opened directly"
 Check Not Forms!frmMajorWorkbench.Visible, "Defaults exposed legacy workbench"
 Check UXGroupNavigate("frmUXBulk", 0), "Return to Members failed"
 Check Forms!frmUXComparison!uxGroupContext = groupID, "Group context lost"
 Check UXGroupNavigate("frmUXComparison", 4), "History navigation failed"
 Check Forms!frmMajorWorkbench!uxMode = "History", "Wrong history mode"
 Check UXGroupNavigate("frmMajorWorkbench", 0), "History to Members failed"
 Check Not Forms!frmMajorWorkbench.Visible, "Members left workbench visible"
 Check UXCloseComparison(), "Return from Members failed"
 Check CurrentProject.AllForms("frmUXGroups").IsLoaded, "Back did not return to groups"
 TestEfficiency = "PASS: " & checks & " targeted refresh, idle timer, duplicate-control cleanup and group navigation assertions"
 Exit Function
Failed:
 TestEfficiency = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function PrepareEventTimer() As Boolean
 On Error GoTo Failed
 timerID = DMin("ReportID", "Report")
 DoCmd.OpenForm "frmEditor", , , , , , CStr(timerID)
 Forms!frmEditor.TimerInterval = 0
 UXSelectSection 0
 Forms!frmEditor!uxFieldSurface.SetFocus
 Forms!frmEditor!uxFieldSurface.Form!txtOfficialName.SetFocus
 Forms!frmEditor!uxFieldSurface.Form!txtOfficialName.Text = "TIMER, SYNTHETIC"
 Check Forms!frmEditor!uxFieldSurface.Form!txtOfficialName.OnChange = "=UXLiveTyped()", "Active textbox OnChange not wired"
 Call Application.Eval(Mid$(Forms!frmEditor!uxFieldSurface.Form!txtOfficialName.OnChange, 2))
 paintBefore = UXLivePaintCount()
 PrepareEventTimer = True
 Exit Function
Failed:
 PrepareEventTimer = False
End Function
Public Function FinishEventTimer() As String
 On Error GoTo Failed
 Check UXLiveTimer(), "Event timer failed"
 Check UXLivePaintCount() = paintBefore + 1, "Dirty timer did not validate exactly once; before=" & paintBefore & "; after=" & UXLivePaintCount()
 Check DLookup("OfficialName", "RecoveryDraft", "ReportID=" & timerID) = "TIMER, SYNTHETIC", "Timer did not checkpoint active typed text"
 paintBefore = UXLivePaintCount(): Call UXLiveTimer()
 Check UXLivePaintCount() = paintBefore, "Clean timer rescanned values"
 FinishEventTimer = "PASS: active OnChange validation and 30-second recovery checkpoint"
 Exit Function
Failed:
 FinishEventTimer = "FAIL: " & Err.Number & " " & Err.Description
End Function

