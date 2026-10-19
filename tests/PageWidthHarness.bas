Attribute VB_Name = "PageWidthHarness"
Option Explicit
Public Function TestPageWidth() As String
 Dim kind As Variant, row As Variant, parts As Variant, f As Access.Form, s As Access.Form, c As Access.Control, id As Long, x As Long, y As Long, w As Long, n As Long, page As Long
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC PAGE WIDTH", "Tester"
 For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
  id = CommitDraft(0, CStr(kind), "Width test", "SYNTHETIC, MEMBER", Null, Null, "")
  DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
  Set f = Forms!frmEditor: Set s = f!uxFieldSurface.Form
  For page = 0 To 1
   If Not UXSelectSection(IIf(page = 0, 0, 5)) Then Err.Raise 5, , "Page open failed"
   For Each row In ReportPageRows(CStr(kind), page)
    parts = Split(CStr(row), "|")
    If parts(0) <> "@guidance" Then
     Set c = s.Controls(UXFieldControl(CStr(parts(0))))
     If c.Left + c.Width > 12200 Then Err.Raise 5, , "Field exceeds paper extent: " & c.Name
     n = n + 1
    End If
   Next
   x = s!txtOfficialName.Left: y = s!txtOfficialName.Top: w = s!txtOfficialName.Width
   UXResizeEditor 28000, 10000
   If f!uxFieldSurface.Width < 24000 Then Err.Raise 5, , "Wide window space not used"
   UXResizeEditor 9000, 7000
   If f!uxFieldSurface.Width < 12600 Or f.ScrollBars <> 3 Or s.ScrollBars <> 3 Then Err.Raise 5, , "Narrow scroll fallback absent"
   If s!txtOfficialName.Left <> x Or s!txtOfficialName.Top <> y Or s!txtOfficialName.Width <> w Then Err.Raise 5, , "Fields reflowed"
   n = n + 3
  Next
  UXLeaveEditor "Discard"
 Next
 TestPageWidth = "PASS: " & n & " right-edge, wide viewport, narrow scrolling and fixed-field assertions": Exit Function
Failed:
 TestPageWidth = "FAIL: " & Err.Number & " " & Err.Description
End Function
