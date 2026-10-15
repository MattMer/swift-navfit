Attribute VB_Name = "ResourceToolsHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8350, , message
 n = n + 1
End Sub
Private Sub RowEntry(ByVal grade As String, ByVal average As Variant, ByVal count As Variant, ByVal verified As Variant)
 With Forms!frmUXSeniorEditor!uxRSCATable.Form
  .Controls("avg" & grade) = average: .Controls("count" & grade) = count: .Controls("date" & grade) = verified
 End With
End Sub
Public Function TestResourceTools() As String
 Dim stage As String, senior As Long, values As Object, f As Access.Form, key As Variant, total As Long, old As Long, db As DAO.Database
 Dim g As Long, other As Long, id As Long, reviewer As Long, template As Long, baseline As Long
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC RESOURCE TOOLS", "Reviewer one": stage = "Senior table"
 Set values = RSSeniorValues(0): values("ReportingSeniorName") = "SYNTHETIC SENIOR": values("ReportingSeniorSSN") = "000000555"
 stage = "Save senior": senior = RSSaveSenior(0, "Synthetic senior", values)
 baseline = RSAddBaseline(senior, "E6", #10/1/2026#, 4, 70, 280, "Synthetic exact history")
 stage = "Open senior": DoCmd.OpenForm "frmUXSeniorEditor", , , , , , CStr(senior): Set f = Forms!frmUXSeniorEditor
 Check Not RSTableDirty(), "Unopened history page reports pending edits"
 RSSelectSeniorPage 3
 Check f!uxRSCATable.Visible And Not f!cboPaygrade.Visible And Not f!cmdAddBaseline.Visible, "Old entry UI still visible"
 Check UBound(RSTableGrades()) = 23, "All paygrades missing"
 For Each key In RSTableGrades()
  Check UXHasControl(f!uxRSCATable.Form, "avg" & key), "Missing RSCA row"
  Check UXHasControl(f!uxRSCATable.Form, "count" & key), "Missing trait count row"
  Check f!uxRSCATable.Form.Controls("date" & key).ShowDatePicker = 1, "Date picker missing"
 Next
 stage = "Loaded dirty state": Check Not RSTableDirty(), "Loaded table dirty"
 stage = "Existing row": Check Nz(f!uxRSCATable.Form!avgE6, -1) = 4 And Nz(f!uxRSCATable.Form!countE6, -1) = 70, "Existing history not loaded"
 total = DCount("*", "RSCABaseline")
 RowEntry "E5", 3.8, 100, #10/14/2026#: RowEntry "O10", 4, 7, #10/14/2026#
 Check RSTableDirty(), "Table edit not detected"
 RSSelectSeniorPage 0: RSSelectSeniorPage 3
 Check f!uxRSCATable.Form!avgE5 = 3.8, "Page switch lost pending cells"
 Check RSSaveTable(), "Table save failed: " & f!lblStatus.Caption
 Check DCount("*", "RSCABaseline") = total + 2, "Wrong number of baseline appends"
 Check Not RSTableDirty(), "Saved table dirty"
 Check RSSaveTable() And DCount("*", "RSCABaseline") = total + 2, "Unchanged rows duplicated"
 RowEntry "E6", 4, 70, #10/14/2026#: Check RSSaveTable(), "Verification date update failed"
 old = DMax("BaselineID", "RSCABaseline", "Paygrade='E6'")
 Check DLookup("TraitSum", "RSCABaseline", "BaselineID=" & old) = 280, "Exact sum lost on date-only update"
 Check DCount("*", "RSCABaseline", "BaselineID=" & baseline) = 1, "Original history overwritten"
 stage = "Validation and rollback": total = DCount("*", "RSCABaseline")
 RowEntry "E1", 4, 7, #10/14/2026#: RowEntry "E2", 8, 7, #10/14/2026#
 Check Not RSSaveTable(), "Invalid RSCA saved"
 Check DCount("*", "RSCABaseline") = total And f!uxRSCATable.Form!avgE1 = 4, "Validation partly saved or lost raw input"
 RowEntry "E2", 4, 7.5, #10/14/2026#: Check Not RSSaveTable(), "Fraction count saved"
 RowEntry "E2", 4, 7, "2026-02-30": Check Not RSSaveTable(), "Invalid date saved"
 RowEntry "E2", 4, 7, #10/14/2026#
 f!lstBaselines.RowSource = "": Set db = CurrentDb: db.TableDefs("RSCABaseline").ValidationRule = "[Paygrade]<>'E2'"
 Check Not RSSaveTable(), "Injected later-row write succeeded"
 Check DCount("*", "RSCABaseline") = total, "Earlier row survived failed transaction"
 db.TableDefs("RSCABaseline").ValidationRule = ""
 Check RSSaveTable(), "Retry failed"
 Check RSViewTableHistory() And f!lstBaselines.Visible, "Prior history inaccessible"
 Check RSViewTableHistory() And f!uxRSCATable.Visible, "Return to table failed"
 RSSelectSeniorPage 0: Check Not f!uxRSCATable.Visible And Not f!lstBaselines.Visible, "History leaked onto detail page"
 Check RSLeaveSeniorEditor(), "Unchanged senior leave failed"
 stage = "Dedicated route templates"
 Check UXOpenRouteTemplates(), "Template destination failed": Set f = Forms!frmUXRouteTemplates
 Check Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded, "Templates opened old group workbench"
 reviewer = AddContributor("Reviewer two")
 f!txtTitle = "Synthetic ordered route": f!txtInstructions = "Synthetic instructions"
 f!cboReviewer = DMin("UserProfileID", "UserProfile"): UXTemplateReviewer False
 f!cboReviewer = reviewer: UXTemplateReviewer False
 Check UXSaveTemplate(), "Template save failed: " & f!lblStatus.Caption
 template = f!lstTemplates
 Check DCount("*", "RouteTemplateStep", "TemplateID=" & template) = 2, "Ordered steps missing"
 Check DLookup("AssigneeID", "RouteTemplateStep", "TemplateID=" & template & " AND StepOrdinal=2") = reviewer, "Reviewer order changed"
 f!txtTitle = "Retained invalid route": Check Not UXSaveTemplate(), "Empty route accepted"
 Check f!txtTitle = "Retained invalid route", "Failed route lost entries"
 DoCmd.Close acForm, "frmUXRouteTemplates", acSaveNo
 stage = "Current group tools"
 g = CreateSummaryGroup("Synthetic E6", "EVAL", "E6")
 other = CreateSummaryGroup("Synthetic officer", "FITREP", "O3")
 Check RSOpenGroupCalculator(g), "Explicit group calculator failed"
 Check CLng(Forms!frmUXGroupRSCA!txtGroupID) = g, "Wrong calculator group"
 Forms!frmUXGroupRSCA!cboSenior = senior: RSGroupSeniorChanged
 Check RSViewGroupSenior(), "Senior record unavailable"
 Check CLng(Forms!frmUXSeniorEditor!txtSeniorID) = senior, "Wrong senior record"
 Check RSLeaveSeniorEditor(), "Senior record leave failed"
 Check RSOpenGroupCalculator(other), "Calculator context switch failed"
 Check CLng(Forms!frmUXGroupRSCA!txtGroupID) = other, "Calculator reused stale group"
 id = CommitDraft(0, "EVAL", "Synthetic", "Synthetic member", Null, Null, "")
 AddGroupMember g, id, "E6"
 DoCmd.OpenForm "frmUXComparison", , , , , , "G:" & g
 Check Forms!frmUXComparison!uxCurrentExport.Visible And Forms!frmUXComparison!uxCurrentSenior.Visible, "Current group tools hidden"
 Check Not UXHasControl(Forms!frmUXComparison, "uxGroupTab3"), "Duplicate comparison tab remains"
 Check Forms!frmUXComparison!uxCurrentSenior.Top > Forms!frmUXComparison!uxTraitGrid.Top + Forms!frmUXComparison!uxTraitGrid.Height, "Group actions precede fields"
 Check UXGroupNavigate("frmUXComparison", 3) And Forms!frmUXComparison!uxGroupTab0.FontBold, "Legacy comparison did not resolve to Members"
 Check UXCurrentGroupTool("frmUXComparison", "Senior"), "Current group senior action failed"
 Check CLng(Forms!frmUXGroupRSCA!txtGroupID) = g, "Senior action used stale context"
 Check UXCurrentGroupTool("frmUXComparison", "Export"), "Current group export failed"
 Check Forms!frmNAVFITExchange!cboGroup = g, "Export lost current group"
 TestResourceTools = "PASS: " & n & " resource assertions": Exit Function
Failed:
 TestResourceTools = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
