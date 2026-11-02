Attribute VB_Name = "PageSwitchHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8955, , message
 n = n + 1
End Sub
Public Function TestPageSwitch() As String
 Dim id As Long, revision As Long, f As Access.Form, s As Access.Form, tag As String, i As Long
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC PAGE SWITCH", "Tester"
 id = CommitDraft(0, "EVAL", "Switch", "SYNTHETIC, MEMBER", Null, Null, "")
 DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
 Set f = Forms!frmEditor: Set s = f!uxFieldSurface.Form
 revision = CurrentRevision(id)
 s!txtOfficialName.SetFocus: s!txtOfficialName.Text = "SWITCH, UNSAVED"
 Check UXSelectSection(5), "Back page failed"
 Check f!txtOfficialName = "SWITCH, UNSAVED" And EditorHasPending(), "Active edit lost on switch"
 ' Save formats names, grades and titles as NAVFIT98A does.
 f!txtRankRate = "it2": Call UXFormatTypedFields
 Check f!txtRankRate = "IT2", "Grade/rate not uppercased at save"
 Check f.Painting And s.Painting, "Painting left disabled"
 Check UXSelectSection(0), "Front page failed"
 Check s!txtOfficialName = "SWITCH, UNSAVED", "Pending field not restored"
 For i = 1 To 8
  Check UXSelectSection(IIf(i Mod 2 = 1, 5, 0)), "Repeated page failed"
  Check f.Painting And s.Painting, "Repeated page disabled painting"
 Next
 Check CurrentRevision(id) = revision, "Page switch created saved revision"
 f.Painting = False: s.Painting = False
 Check UXSelectSection(5), "Nested no-paint switch failed"
 Check Not f.Painting And Not s.Painting, "Caller painting state not retained"
 f.Painting = True: s.Painting = True
 tag = f!uxReport.Tag: f!uxReport.Tag = "LOCAL:invalid"
 Check Not UXSelectSection(0), "Injected layout error did not fail"
 Check f.Painting And s.Painting, "Layout error stranded painting"
 f!uxReport.Tag = tag
 Check UXSelectSection(0), "Could not switch after layout error"
 Check s!txtOfficialName = "SWITCH, UNSAVED", "Failure discarded pending edit"
 Check UXLeaveEditor("Discard"), "Cleanup failed"
 TestPageSwitch = "PASS: " & n & " page switch, active-text preservation and paint restoration assertions": Exit Function
Failed:
 TestPageSwitch = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
