Attribute VB_Name = "PageTimingHarness"
Option Explicit
Public Function MeasurePages() As String
 Dim id As Long, i As Long, t As Double, values As String
 ConfigureWorkspace "SYNTHETIC PAGE TIMING", "Tester"
 id = CommitDraft(0, "EVAL", "Timing", "SYNTHETIC, MEMBER", Null, Null, "")
 DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
 For i = 1 To 10
  t = Timer: UXSelectSection IIf(i Mod 2 = 1, 5, 0)
  If Len(values) > 0 Then values = values & ","
  values = values & Format$((Timer - t) * 1000, "0.000")
 Next
 UXLeaveEditor "Discard"
 MeasurePages = values
End Function

Public Function MeasureExistingPages() As String
 Dim id As Long, i As Long, t As Double, values As String
 id = DMin("ReportID", "Report", "FormKind='FITREP' AND IsArchived=False")
 DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
 For i = 1 To 10
  t = Timer: UXSelectSection IIf(i Mod 2 = 1, 5, 0)
  If Len(values) > 0 Then values = values & ","
  values = values & Format$((Timer - t) * 1000, "0.000")
 Next
 UXLeaveEditor "Discard"
 MeasureExistingPages = values
End Function