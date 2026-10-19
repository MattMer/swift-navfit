Attribute VB_Name = "TableSortHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8410, , message
 n = n + 1
End Sub
Public Function TestTableSort() As String
 Dim c As Access.Control, db As DAO.Database, id As Long, other As Long, g As Long, before As Long, f As Access.Form, message As String, senior As Long, values As Object
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC SORT", "Tester"
 id = CommitDraft(0, "EVAL", "Z report", "ZULU, MEMBER", #1/1/2026#, #9/30/2026#, "")
 other = CommitDraft(0, "EVAL", "A report", "ALPHA, MEMBER", #1/1/2026#, #9/30/2026#, "")
 DoCmd.OpenForm "frmLibrary"
 Set c = Forms!frmLibrary!lstReports
 c.ColumnCount = 4: c.ColumnWidths = "0;2500;2000;2000": c.ColumnHeads = True
 c.RowSourceType = "Table/Query"
 c.RowSource = "SELECT ReportID,MemberName,PeriodEnd,Title FROM Report WHERE ReportID IN (" & id & "," & other & ") ORDER BY ReportID"
 c.Value = id
 Check UXSortList("frmLibrary", "lstReports", 1), "Query sort failed: " & c.ControlTipText
 Check c.Column(1, 1) = "ALPHA, MEMBER", "Text order ascending wrong"
 Check c.Value = id, "Selected report lost"
 Check UXSortList("frmLibrary", "lstReports", 1), "Reverse query sort failed"
 Check c.Column(1, 1) = "ZULU, MEMBER", "Text order descending wrong"
 Check c.ListCount = 3, "Query filter lost or heading consumed data"
 'Changing the source filter must not resurrect the earlier cached query.
 c.RowSource = "SELECT ReportID,MemberName,PeriodEnd,Title FROM Report WHERE ReportID=" & other
 Check UXSortList("frmLibrary", "lstReports", 2), "Filtered date sort failed"
 Check c.ListCount = 2 And c.Column(0, 1) = other, "New filter lost"
 Check Not UXSortNativeHeader("frmLibrary", "lstReports", 100, 1500), "Body click triggered sorting"
 Check UXSortNativeHeader("frmLibrary", "lstReports", 100, 100), "Native header click did not sort"
 c.ColumnCount = 3: c.ColumnHeads = False: c.RowSourceType = "Value List"
 c.RowSource = "1;""Z; quoted"";10;2;Alpha;2;3;Beta;1": c.Value = 1
 Check UXSortList("frmLibrary", "lstReports", 2), "Numeric Value List sort failed"
 Check c.Column(0, 0) = "3" And c.Column(2, 2) = "10", "Numbers sorted as text: " & c.RowSource
 Check c.Value = 1 And c.Column(1, 2) = "Z; quoted", "Quoted cell or selection lost"
 c.ColumnHeads = True: c.RowSource = "ID;Member;Mark;1;Z;4;2;A;3"
 Check UXSortList("frmLibrary", "lstReports", 1), "Heading Value List sort failed"
 Check c.Column(1, 0) = "Member" And c.Column(1, 1) = "A", "Heading row lost"
 c.ColumnHeads = False: c.RowSource = ""
 Check UXSortList("frmLibrary", "lstReports", 1), "Empty sort failed"
 DoCmd.Close acForm, "frmLibrary", acSaveNo
 g = CreateSummaryGroup("SYNTHETIC SORT GROUP", "EVAL", "E5")
 AddGroupMember g, id, "E5": AddGroupMember g, other, "E5"
 DoCmd.OpenForm "frmUXComparison", , , , , , "G:" & g
 before = DCount("*", "ReportRevision")
 Check UXSortGrid("Member"), "Grid sort failed: " & Forms!frmUXComparison!lblStatus.Caption
 Set f = Forms!frmUXComparison!uxTraitGrid.Form
 f.Recordset.MoveFirst
 Check f!ReportID = other, "Grid row identity sorted incorrectly"
 Check UXSortGrid("Member"), "Reverse grid sort failed"
 f.Recordset.MoveFirst
 Check f!ReportID = id, "Grid reverse order wrong"
 Check DCount("*", "ReportRevision") = before, "Sorting created report revision"
 f!Trait33 = "4.0": Check UXGridQueue("Trait33"), "Sorted cell did not queue"
 Check UXSortGrid("Trait33"), "Sorting with pending cell failed"
 Check DLookup("Trait33", "Report", "ReportID=" & id) = "4.0", "Sorted row saved to wrong report"
 Check IsNull(DLookup("Trait33", "Report", "ReportID=" & other)), "Sorted edit changed another report"
 Check UXCloseComparison(False), "Grid failed to close"
 DoCmd.OpenForm "frmUXNativeGroups"
 Set c = Forms!frmUXNativeGroups!lstGroups
 c.RowSource = "1;Zulu;EVAL;E5;10;2;Alpha;EVAL;E5;2"
 c.Selected(0) = True: c.Selected(1) = True
 Check UXSortList("frmUXNativeGroups", "lstGroups", 4), "Multi-selection sort failed"
 Check c.Column(0, 0) = "2" And c.Selected(0) And c.Selected(1), "Multi-selection or numeric group count lost"
 DoCmd.Close acForm, "frmUXNativeGroups", acSaveNo
 Set values = RSSeniorValues(0): values("ReportingSeniorName") = "SYNTHETIC RS"
 senior = RSSaveSenior(0, "Synthetic sorting senior", values)
 DoCmd.OpenForm "frmUXSeniorEditor", , , , , , CStr(senior)
 RSSelectSeniorPage 3
 Set f = Forms!frmUXSeniorEditor!uxRSCATable.Form
 f!avgE5 = 4.5: f!countE5 = 100: f!dateE5 = #10/1/2026#
 f!avgE6 = 3.5: f!countE6 = 200: f!dateE6 = #9/1/2026#
 before = DCount("*", "RSCABaseline")
 Check RSTableDirty(), "Table entry not pending"
 Check RSTableSort("avg"), "RSCA numeric sort failed"
 Check f!avgE6.Top < f!avgE5.Top, "RSCA order wrong"
 Check f!avgE5 = 4.5 And f!countE5 = 100 And f!dateE5 = #10/1/2026#, "RSCA pending values lost"
 Check RSTableDirty() And DCount("*", "RSCABaseline") = before, "RSCA sort saved data"
 Check RSTableSort("avg") And f!avgE5.Top < f!avgE6.Top, "RSCA reverse order wrong"
 Check f!avgE5.TabIndex < f!avgE6.TabIndex, "RSCA keyboard order did not follow rows"
 RSTableLoad
 DoCmd.Close acForm, "frmUXSeniorEditor", acSaveNo
 DoCmd.OpenForm "frmUXNavyWriter"
 Check UXSortHeaders(Forms!frmUXNavyWriter), "Custom heading layout failed"
 Check Forms!frmUXNavyWriter!sort_lstCatalog_1.Visible And Forms!frmUXNavyWriter!sort_lstCatalog_1.Caption = "Rate", "Narrative rate heading missing"
 Check Forms!frmUXNavyWriter!sort_lstCatalog_3.Caption = "Example", "Narrative example heading wrong"
 DoCmd.Close acForm, "frmUXNavyWriter", acSaveNo
 TestTableSort = "PASS: " & n & " query/value-list/header/selection/filter/grid sort assertions": Exit Function
Failed:
 message = Err.Number & " " & Err.Description
 TestTableSort = "FAIL after " & n & ": " & message
End Function
