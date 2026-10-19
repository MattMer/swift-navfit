Attribute VB_Name = "VisualHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8990, , message
 n = n + 1
End Sub
Public Function TestVisual() As String
 Dim name As Variant, f As Access.Form, id As Long, before As Long, x As Long, y As Long, w As Long, i As Long
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC VISUAL TEST", "Tester"
 For Each name In Array("frmHome", "frmLibrary", "frmUXGroups", "frmMyWork", "frmUXResources", "frmUXSettings")
  DoCmd.OpenForm CStr(name): Set f = Forms(CStr(name))
  Check UXApplyVisual(f), "Palette failed: " & name
  For i = 0 To 5
   Check f.Controls("uxNav" & i).BackColor = UXColor(IIf(i = nPrimary(CStr(name)), "Active", "Rail")), "Wrong navigation emphasis: " & name
   Check f.Controls("uxNav" & i).ForeColor = RGB(255, 255, 255), "Navigation contrast missing"
  Next
  Check f!cmdWindowHelp.BackColor = UXColor("Help"), "Help palette missing"
  Check UXResizeWorkspace(CStr(name), 9000, 7000), "Compact resize failed"
  Check f!cmdWindowHelp.ForeColor = UXColor("HelpInk"), "Resize lost help palette"
  DoCmd.Close acForm, CStr(name), acSaveNo
 Next
 id = CommitDraft(0, "EVAL", "Visual", "SYNTHETIC, MEMBER", Null, Null, "")
 DoCmd.OpenForm "frmLibrary"
 Forms!frmLibrary!lstReports = id
 Check EditSelected(), "Report list open binding failed"
 Set f = Forms!frmEditor
 f.TimerInterval = 0
 Check UXWindowHelp("frmEditor"), "Editor help missing"
 Check Forms!frmUXHelp!lstTopics.Column(1, 1) = "Create and save reports", "Help topic truncated"
 CloseHelp
 x = f!uxFieldSurface.Left: y = f!uxFieldSurface.Top: w = f!uxFieldSurface.Width
 before = DCount("*", "ReportRevision")
 f!uxDirty.Caption = "Unsaved changes | 2 issues": Check UXApplyVisual(f), "Pending palette failed"
 Check f!uxDirty.BackColor = UXColor("Warning"), "Pending not amber: " & f!uxDirty.Caption & " / " & f!uxDirty.BackColor
 f!uxError.Caption = "Not saved - retry": UXApplyVisual f
 Check f!uxError.BackColor = UXColor("Error"), "Failure not red"
 f!lblSaved.Caption = "Saved revision 1 - local draft": UXApplyVisual f
 Check f!lblSaved.BackColor = UXColor("Success"), "Saved not green"
 f!lblRecovery.Caption = "Recovery: no unsaved edits to checkpoint.": UXApplyVisual f
 Check f!lblRecovery.BackColor = UXColor("Success"), "Clean checkpoint misclassified"
 Check DCount("*", "ReportRevision") = before, "Palette created a revision"
 Check f!uxFieldSurface.Left = x And f!uxFieldSurface.Top = y And f!uxFieldSurface.Width = w, "Palette changed fixed geometry"
 Check UXLeaveEditor("Discard"), "Editor cleanup failed"
 TestVisual = "PASS: " & n & " semantic palette, navigation, status and fixed-geometry assertions": Exit Function
Failed:
 TestVisual = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
Private Function nPrimary(ByVal name As String) As Long
 Select Case name
  Case "frmLibrary": nPrimary = 1
  Case "frmUXGroups": nPrimary = 2
  Case "frmMyWork": nPrimary = 3
  Case "frmUXResources": nPrimary = 4
  Case "frmUXSettings": nPrimary = 5
 End Select
End Function
