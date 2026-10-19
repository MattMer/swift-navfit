Attribute VB_Name = "CompatPagesHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8400, , message
 n = n + 1
End Sub
Public Function TestCompatPages() As String
 Dim kind As Variant, id As Long, f As Access.Form, s As Access.Form, row As Variant, key As String, page As Long, before As Long, renders As Long, x As Long, y As Long, width As Long, groupID As Long, values As Object
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC FIXED PAGES", "Tester"
 For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
  id = CommitDraft(0, CStr(kind), "Synthetic fixed page", "SYNTHETIC, MEMBER", #1/1/2026#, #9/30/2026#, "", administration:=Array("SYNTHETIC, MEMBER", "", "", "", ""))
  DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
  Set f = Forms!frmEditor: Set s = f!uxFieldSurface.Form
  Check UXFixedEnabled(), "Fixed report pages unavailable"
  Check f!uxSection.Caption = "Front Page", "Front page failed: " & f!uxError.Caption
  For page = 0 To 1
   Check UXSelectSection(IIf(page = 0, 0, 5)), "Page could not open: " & f!uxError.Caption
   For Each row In ReportPageRows(CStr(kind), page)
    key = Split(CStr(row), "|")(0)
    If key <> "@guidance" Then Check s.Controls(UXFieldControl(key)).Visible, "Field missing on correct page: " & kind & " " & key
   Next
  Next
  For Each row In Array("Trait33", "Trait34", "Trait35", "Trait36", "Trait37", "Trait38", "Trait39")
   Check ReportPageForField(CStr(kind), CStr(row)) = 0, "Trait not on Front Page: " & row
  Next
  UXSelectSection 0
  s!txtOfficialName = "PENDING, MEMBER": UXSurfaceEdited "OfficialName"
  before = DCount("*", "ReportRevision")
  UXSelectSection 5: UXSelectSection 0
  Check s!txtOfficialName = "PENDING, MEMBER" And EditorHasPending(), "Page switch lost pending value"
  Check DCount("*", "ReportRevision") = before, "Page switch saved revision"
  x = s!txtOfficialName.Left: y = s!txtOfficialName.Top: width = s!txtOfficialName.Width: renders = UXFixedRenderCount()
  UXResizeEditor 16000, 10000: UXResizeEditor 28000, 14500: UXResizeEditor 12000, 8000
  Check s!txtOfficialName.Left = x And s!txtOfficialName.Top = y And s!txtOfficialName.Width = width, "Window change reflowed paper fields"
  Check UXFixedRenderCount() = renders, "Window change rebuilt fields"
  Check f!uxRail.Visible And f!uxNav0.Visible And f!uxNav5.Visible, "Main navigation rail unavailable"
  Check f!uxFieldSurface.Left > f!uxRail.Left + f!uxRail.Width, "Report overlaps the main navigation rail"
  Check f!cmdSave.Left > f!uxRail.Left + f!uxRail.Width And f!uxErrorList.Left > f!uxRail.Left + f!uxRail.Width, "Report actions/errors overlap navigation"
  Check UXFixedOpenField("Trait39") And s!cboTrait39.Visible And f!uxSection.Caption = "Front Page", "Validation could not navigate to front-page trait"
  Check UXSelectSection(7) And s!uxNoteList.Visible, "Local notes unavailable"
  Check s!sort_uxNoteDrafts_1.Visible, "Pending note table sorting headings unavailable"
  Check UXSelectSection(8) And f!cmdRoute.Visible, "Local review unavailable"
  Check UXSelectSection(9) And f!lstRevisions.Visible, "History unavailable"
  Check UXLeaveEditor("Discard"), "Unable to close unchanged/explicitly discarded report"
 Next
 groupID = CreateSummaryGroup("Closeout group", "EVAL", "E5")
 id = CommitDraft(0, "EVAL", "Closeout test", "SYNTHETIC, MEMBER", #1/1/2026#, #9/30/2026#, "")
 AddGroupMember groupID, id, "E5"
 Check GroupCloseout(groupID, 1, 1, #9/30/2026#, #9/30/2026#) = "2026-09-30", "Closeout date wrong"
 Check GroupCloseout(groupID, 2, 2, #8/31/2026#, #9/30/2026#) = "Mixed", "Mixed end dates hidden"
 Check GroupCloseout(groupID, 2, 1, #9/30/2026#, #9/30/2026#) = "Incomplete", "Missing end date hidden"
 DoCmd.OpenForm "frmUXGroups": Check UXGroupList(), "Group list query failed"
 Check Forms!frmUXGroups!lstGroups.ColumnCount = 8, "Closeout column absent"
 Check UXSortList("frmUXGroups", "lstGroups", 7), "Closeout date heading sort failed"
 TestCompatPages = "PASS: " & n & " fixed report-page, pending buffer, local workflow and closeout assertions": Exit Function
Failed:
 TestCompatPages = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
