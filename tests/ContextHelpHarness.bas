Attribute VB_Name = "ContextHelpHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8900, , message
 n = n + 1
End Sub
Public Function TestContextHelp() As String
 Dim id As Long, before As Long, f As Access.Form, name As Variant, x As Long, y As Long, w As Long, item As AccessObject, hasLegacy As Boolean
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC CONTEXT HELP", "Tester"
 id = CommitDraft(0, "EVAL", "Help test", "SYNTHETIC, MEMBER", #1/1/2026#, #9/30/2026#, "")
 DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
 Set f = Forms!frmEditor
 f!txtOfficialComments = "Pending help preservation": UXEditorChanged
 before = DCount("*", "ReportRevision")
 x = f!uxFieldSurface.Left: y = f!uxFieldSurface.Top: w = f!uxFieldSurface.Width
 Check UXWindowHelp("frmEditor"), "Editor help failed"
 Check Forms!frmUXHelp!lstTopics = 1, "Editor topic incorrect"
 Check InStr(Forms!frmUXHelp!txtGuide, "Report editor:") = 1, "Window-specific introduction missing"
 Check f!txtOfficialComments = "Pending help preservation" And EditorHasPending(), "Help changed pending edits"
 Check DCount("*", "ReportRevision") = before, "Help created a revision"
 Check CurrentProject.AllForms("frmEditor").IsLoaded, "Help navigated away from editor"
 Check UXWindowHelpLayout("frmEditor"), "Editor help placement failed"
 Check f!cmdWindowHelp.Top = 150 And f!cmdWindowHelp.Left + f!cmdWindowHelp.Width <= f.InsideWidth, "Help not at top right"
 Check f!uxFieldSurface.Left = x And f!uxFieldSurface.Top = y And f!uxFieldSurface.Width = w, "Help moved report fields"
 Check ShowHelp("Saving") And Forms!frmUXHelp!lstTopics = 1, "Legacy save help did not consolidate"
 Check ShowHelp("Folders") And Forms!frmUXHelp!lstTopics = 2, "Legacy folder help did not consolidate"
 Check OpenScreen("frmHelp"), "Bare legacy route failed"
 For Each item In CurrentProject.AllForms
  If item.Name = "frmHelp" Then hasLegacy = True
 Next
 Check Not hasLegacy And Screen.ActiveForm.Name = "frmUXHelp", "Legacy help window retained or wrong guide opened"
 CloseHelp
 Check EditorHasPending(), "Closing help discarded edits"
 Check UXLeaveEditor("Discard"), "Test cleanup failed"
 For Each name In Array("frmHome", "frmLibrary", "frmUXGroups", "frmMyWork", "frmUXSettings", "frmUXResources")
  DoCmd.OpenForm CStr(name)
  Set f = Forms(CStr(name))
  Check UXHasControl(f, "cmdWindowHelp"), "Help button missing: " & name
  Check UXWindowHelp(CStr(name)), "Context help failed: " & name
  Check Forms!frmUXHelp!lstTopics = UXWindowHelpTopic(CStr(name)), "Wrong topic: " & name
  Check InStr(Forms!frmUXHelp!lblGuide.Caption, f.Caption) > 0, "Help heading loses originating window"
  If UXHasControl(f, "uxCompactNav") Then
   Check UXResizeWorkspace(CStr(name), 9000, 7000), "Compact resize failed"
   Check f!cmdWindowHelp.Left = f.InsideWidth - 800, "Compact help did not follow resize"
   Check f!cmdWindowHelp.Left > 0 And f!cmdWindowHelp.Left + 500 <= f.InsideWidth, "Compact help clipped"
  End If
  CloseHelp: DoCmd.Close acForm, CStr(name), acSaveNo
 Next
 TestContextHelp = "PASS: " & n & " context help, buffer preservation and layout assertions": Exit Function
Failed:
 TestContextHelp = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
