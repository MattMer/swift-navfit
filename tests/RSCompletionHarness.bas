Attribute VB_Name = "RSCompletionHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 4860, , message
 n = n + 1
End Sub
Private Function Member(ByVal senior As Object, ByVal name As String) As Long
 Dim f As Object, key As Variant
 Member = CommitDraft(0, "EVAL", name, name, #1/1/2026#, #9/30/2026#, "", "Test")
 Set f = CreateObject("Scripting.Dictionary")
 For Each key In senior.Keys: f.Add CStr(key), senior(key): Next
 f.Add "ObservationStatus", "Observed"
 For Each key In TraitKeys(): f.Add CStr(key), "4.0": Next
 CommitSavedFields Member, CurrentRevision(Member), f, "Synthetic accounting fixture"
End Function
Private Function RejectComplete(ByVal id As Long) As Boolean
 On Error GoTo Rejected
 SetGroupWorkflow id, GroupStamp(id), "Completed"
 Exit Function
Rejected:
 RejectComplete = True
End Function
Private Function RejectSQL(ByVal sql As String) As Boolean
 On Error GoTo Rejected
 CurrentDb.Execute sql, dbFailOnError
 Exit Function
Rejected:
 RejectSQL = True
End Function
Private Function RejectSenior(ByVal id As Long, ByVal senior As Long) As Boolean
 On Error GoTo Rejected
 RSSetGroupBaseline id, senior, 0, False
 Exit Function
Rejected:
 RejectSenior = True
End Function
Private Function RejectExport(ByVal senior As Long, ByVal path As String) As Boolean
 On Error GoTo Rejected
 RSExportHistory senior, path
 Exit Function
Rejected:
 RejectExport = True
End Function
Private Sub Counts(ByVal senior As Long, ByVal count As Long, ByVal total As Double)
 Dim id As Long
 id = RSLatestBaseline(senior, "E6")
 Check CLng(DLookup("TraitCount", "RSCABaseline", "BaselineID=" & id)) = count, "Wrong running trait count"
 Check Abs(CDbl(DLookup("TraitSum", "RSCABaseline", "BaselineID=" & id)) - total) < 0.00001, "Wrong running trait sum"
End Sub
Public Function TestRSCompletion() As String
 Dim senior As Object, sid As Long, b As Long, g As Long, h As Long, duplicate As Long, a As Long, c As Long, key As Variant, f As Object, p As Object, control As Access.Control, entries As Long, rev As Long, nob As Long, estimated As Long, late As Long, original As String
 Dim posting As Long, legacy As Long, legacyReport As Long, legacyBaseline As Long, xml As String, sid2 As Long
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC RSCA COMPLETION", "Tester"
 Set senior = CreateObject("Scripting.Dictionary")
 For Each key In RSSeniorKeys(): senior.Add CStr(key), Null: Next
 senior("ReportingSeniorSSN") = "000000001": senior("ReportingSeniorName") = "SYNTHETIC"
 sid = RSSaveSenior(0, "Synthetic senior", senior)
 b = RSAddBaseline(sid, "E6", #12/31/2025#, 4.2, 100, 420, "Synthetic BOL observation")
 a = Member(senior, "Member A"): c = Member(senior, "Member B")
 g = CreateSummaryGroup("First", "EVAL", "E6"): AddGroupMember g, a, "E6"
 Check RejectComplete(g), "Missing baseline allowed completion"
 Check GroupWorkflow(g) = "Initialized" And DCount("*", "RSCABaseline") = 1, "Failed completion was not atomic"
 RSSetGroupBaseline g, sid, b, True
 Set p = RSGroupPSG(g): Check Abs(p("Value") - 448# / 107) < 0.00001, "Initial projection incorrect"
 'The current RSCA (before this group) shows beside the PSG RSCA with the change.
 Check Abs(p("Current") - 4.2) < 0.00001 And UXRSCADeltaText(p) = "Current RSCA: 4.20 | PSG RSCA: 4.19 (-0.01)", "Current RSCA delta wrong: " & UXRSCADeltaText(p)
 Check UXRSCADelta(0.004) = "0.00" And UXRSCADelta(-0.004) = "0.00" And UXRSCADelta(0.05) = "+0.05" And UXRSCADelta(-0.25) = "-0.25", "RSCA delta sign wrong"
 rev = CurrentRevision(a)
 SetGroupWorkflow g, GroupStamp(g), "Completed": Counts sid, 107, 448
 Check GroupWorkflow(g) = "Completed" And CurrentRevision(a) = rev, "Completion changed report revision"
 posting = RSPosting(g)
 Check posting > 0 And DLookup("RevisionID", "GroupCompletionReport", "CompletionID=" & posting) = rev, "Posting did not freeze the counted revision"
 Check InStr(Nz(DLookup("KeySpec", "SummaryGroup", "GroupID=" & g), ""), "RSCAPostedBaseline") = 0, "New posting still writes XML accounting"
 CurrentDb.Execute "UPDATE Report SET Trait33='5.0' WHERE ReportID=" & a, dbFailOnError
 Check RejectComplete(g), "Tampered trait cache was posted instead of the saved revision"
 CurrentDb.Execute "UPDATE Report SET Trait33='4.0' WHERE ReportID=" & a, dbFailOnError
 Check RejectSQL("DELETE FROM RSCABaseline WHERE BaselineID=" & b), "Posted ancestor deletion was accepted"
 Check DLookup("TraitCount", "RSCABaseline", "BaselineID=" & b) = 100, "Original baseline overwritten"
 entries = DCount("*", "RSCABaseline")
 SetGroupWorkflow g, GroupStamp(g), "Completed"
 Check DCount("*", "RSCABaseline") = entries, "Repeat completion double counted"
 Set p = RSGroupPSG(g): Check Abs(p("Value") - 448# / 107) < 0.00001, "Completed group counted twice in projection"
 h = CreateSummaryGroup("Second", "EVAL", "E6"): AddGroupMember h, c, "E6": RSSetGroupBaseline h, sid, b, True
 Set p = RSGroupPSG(h): Check Abs(p("Value") - 476# / 114) < 0.00001, "Projection omitted earlier completed group"
 Check Abs(p("Current") - 448# / 107) < 0.00001 And InStr(UXRSCADeltaText(p), "Current RSCA: 4.19 | PSG RSCA: 4.18 (-0.01)") = 1, "Current RSCA omitted the completed group"
 SetGroupWorkflow h, GroupStamp(h), "Completed": Counts sid, 114, 476
 SetGroupWorkflow g, GroupStamp(g), "In-Work": Counts sid, 114, 476
 Set f = CreateObject("Scripting.Dictionary")
 For Each key In TraitKeys(): f.Add CStr(key), "3.0": Next
 CommitSavedFields a, CurrentRevision(a), f, "Correct completed traits"
 SetGroupWorkflow g, GroupStamp(g), "Completed": Counts sid, 114, 469
 Check DLookup("PriorCompletionID", "GroupCompletion", "CompletionID=" & RSPosting(g)) = posting, "Correction lost prior posting history"
 Check DLookup("RevisionID", "GroupCompletionReport", "CompletionID=" & posting) = rev, "Correction rewrote frozen revision"
 Check Left$(RSReconcileAccounting(), 5) = "PASS:", "Relational accounting does not reconcile"
 'The RSCA history names each group posted on top of the BOL baseline.
 Dim summary As String, trail As String
 trail = RSHistoryList(sid, "E6", summary)
 Check InStr(trail, "Group completed: First") > 0 And InStr(trail, "Group completed: Second") > 0 And InStr(trail, "Group re-completed (correction): First") > 0 And InStr(trail, "Earlier:") = 0, "RSCA history did not name the posting groups"
 Check summary = "E6 BOL baseline 2025-12-31: RSCA 4.20 (100 traits); now 4.11 (114 traits) after 3 group completions.", "RSCA history summary wrong: " & summary
 Check RSBaselineSource(RSLatestBaseline(sid, "E6")) = "BOL 2025-12-31 + 3 groups", "Paygrade table source wrong"
 'The RSCA history export: the BOL baseline, each counted report (a correction replaces
 'the group's first posting), reports in open groups (listed, not counted), then the calculation.
 Dim d As Long, o As Long, rows As Collection, v As Variant, found As Long, path As String, x As DAO.Database, xr As DAO.Recordset, t As DAO.TableDef, sheet As String, readRows As Long
 d = Member(senior, "Member D"): o = CreateSummaryGroup("Open", "EVAL", "E6"): AddGroupMember o, d, "E6": RSSetGroupBaseline o, sid, b, True
 Set rows = RSHistoryExportRows(sid)
 Check rows(1)(1) = "Reporting senior" And rows(1)(4) = "Synthetic senior", "Export heading row wrong"
 For Each v In rows
  If Nz(v(0), "") = "E6" Then
   Select Case v(1)
    Case "BOL baseline": Check v(2) = #12/31/2025# And v(7) = 100 And v(8) = 420 And v(9) = 4.2, "Export baseline row wrong": found = found + 1
    Case "Counted report"
     If v(4) = "Member A" Then Check v(3) = "First" And v(6) = 3 And v(7) = 7 And v(8) = 21, "Export corrected report wrong": found = found + 1
     If v(4) = "Member B" Then Check v(3) = "Second" And v(6) = 4 And v(7) = 7 And v(8) = 28, "Export counted report wrong": found = found + 1
     Check v(4) <> "Member D", "Open group report was counted"
    Case "Not counted: group not completed": Check v(4) = "Member D" And v(3) = "Open" And v(6) = 4 And v(7) = 7, "Export open report wrong": found = found + 1
    Case "RSCA calculation": Check v(7) = 114 And v(8) = 469 And v(9) = 4.11 And InStr(v(10), "(420.00 baseline + 49.00 counted reports) / (100 + 14 traits)") = 1 And InStr(v(10), "differs") = 0, "Export calculation wrong: " & v(10): found = found + 1
    Case "RSCA if open groups complete": Check v(7) = 121 And v(8) = 497 And v(9) = 4.11, "Export projection wrong": found = found + 1
   End Select
  End If
 Next
 Check found = 6, "Export rows missing: " & found
 path = CurrentProject.Path & "\rs-history-" & Format$(Now, "hhnnss") & ".xlsx"
 Check RSExportHistory(sid, path) = rows.Count, "Export row count wrong"
 Set x = DBEngine.OpenDatabase(path, False, True, "Excel 12.0 Xml;HDR=YES;IMEX=1;")
 For Each t In x.TableDefs
  If Len(sheet) = 0 And Right$(Replace(t.Name, "'", ""), 1) = "$" Then sheet = Replace(t.Name, "'", "")
 Next
 Set xr = x.OpenRecordset("SELECT * FROM [" & sheet & "]", dbOpenSnapshot)
 Check xr.Fields(1).Name = "Entry" And xr.Fields(6).Name = "ITA" And xr.Fields(9).Name = "RSCA", "Workbook headings wrong"
 found = 0
 Do While Not xr.EOF
  readRows = readRows + 1
  If Nz(xr.Fields(1).Value, "") = "RSCA calculation" And Nz(xr.Fields(0).Value, "") = "E6" Then
   If Val(Nz(xr.Fields(9).Value, 0)) = 4.11 And Val(Nz(xr.Fields(7).Value, 0)) = 114 Then found = found + 1
  End If
  xr.MoveNext
 Loop
 xr.Close: x.Close
 Check readRows = rows.Count And found = 1, "Workbook rows differ: " & readRows & " of " & rows.Count
 Check RejectExport(sid, path), "Export replaced an existing workbook"
 Kill path
 Check RejectExport(sid, Left$(path, Len(path) - 5) & ".csv"), "Export wrote a non-workbook file"
 'The editor shows the same trail as a list and opens the selected group.
 DoCmd.OpenForm "frmUXSeniorEditor", , , , , , CStr(sid)
 RSSelectSeniorPage 3: Check RSViewTableHistory(), "RSCA history view failed"
 With Forms!frmUXSeniorEditor
  Check !lstBaselines.Visible And !cboHistoryGrade = "E6" And !lstBaselines.ListCount = 5, "RSCA history list did not show the trail"
  Check !lstBaselines.Column(2, 1) = "BOL baseline (entered)" And !lstBaselines.Column(2, 2) = "Group completed: First" And !lstBaselines.Column(4, 2) = "7" And !lstBaselines.Column(6, 4) = "4.11", "RSCA history rows wrong"
  Check InStr(!uxRSCATable.Form!srcE6.Caption, "+ 3 groups") > 0, "Paygrade table source column missing"
  Check !cmdExportRSHistory.Visible And !cmdExportRSHistory.OnClick = "=RSExportHistoryUI()", "RSCA history export button missing"
  !lstBaselines = CStr(h)
 End With
 Check RSHistoryOpenGroup() And CurrentProject.AllForms("frmUXComparison").IsLoaded And Not CurrentProject.AllForms("frmUXSeniorEditor").IsLoaded, "Open selected group failed"
 DoCmd.Close acForm, "frmUXComparison", acSaveNo
 sid2 = RSSaveSenior(0, "Second synthetic senior", senior)
 Check RejectSenior(g, sid2), "Counted group reassigned to another senior"
 Check DCount("*", "RSCABaseline") = entries + 2, "Correction did not append one delta entry"
 duplicate = CreateSummaryGroup("Duplicate", "EVAL", "E6"): AddGroupMember duplicate, a, "E6": RSSetGroupBaseline duplicate, sid, b, True
 entries = DCount("*", "RSCABaseline")
 Check RejectComplete(duplicate), "Same report counted through two groups"
 Check GroupWorkflow(duplicate) = "Initialized" And DCount("*", "RSCABaseline") = entries, "Duplicate rejection changed totals/status"
 f.RemoveAll: f.Add "Trait33", Null: CommitSavedFields c, CurrentRevision(c), f, "Incomplete traits"
 Check RejectComplete(h), "Incomplete observed traits allowed completion"
 Check DCount("*", "RSCABaseline") = entries, "Failed trait validation appended totals"
 'Force the final status write to fail AFTER the new baseline is appended.
 late = CreateSummaryGroup("Late rollback", "EVAL", "E6")
 nob = Member(senior, "Member C"): AddGroupMember late, nob, "E6": RSSetGroupBaseline late, sid, b, True
 CurrentDb.TableDefs("SummaryGroup").Fields("VersionCounter").ValidationRule = "< " & (CLng(DLookup("VersionCounter", "SummaryGroup", "GroupID=" & late)) + 1)
 Check RejectComplete(late), "Injected late status failure accepted"
 CurrentDb.TableDefs("SummaryGroup").Fields("VersionCounter").ValidationRule = ""
 Check DCount("*", "RSCABaseline") = entries And GroupWorkflow(late) = "Initialized" And Len(GroupPolicyValue(late, "RSCAPostedBaseline")) = 0, "Late failure retained baseline or marker"
 Check RSPosting(late) = 0 And DCount("*", "GroupCompletion", "GroupID=" & late) = 0, "Late failure retained relational posting"
 'Whole-report NOB adds no graded traits.
 f.RemoveAll: f.Add "ObservationStatus", "NOB": CommitSavedFields nob, CurrentRevision(nob), f, "Whole NOB"
 SetGroupWorkflow late, GroupStamp(late), "Completed": Counts sid, 114, 469
 'Preserve estimated provenance when no exact historic sum was supplied.
 estimated = CreateSummaryGroup("Estimated", "EVAL", "E5")
 nob = Member(senior, "Estimated member"): AddGroupMember estimated, nob, "E5"
 b = RSAddBaseline(sid, "E5", #12/31/2025#, 4.21, 100, Null, "Rounded synthetic observation")
 RSSetGroupBaseline estimated, sid, b, True
 SetGroupWorkflow estimated, GroupStamp(estimated), "Completed"
 b = RSLatestBaseline(sid, "E5")
 Check DLookup("TraitCount", "RSCABaseline", "BaselineID=" & b) = 107 And IsNull(DLookup("TraitSum", "RSCABaseline", "BaselineID=" & b)), "Estimated history became falsely exact"
 Set p = RSGroupPSG(estimated): Check p("Estimated") And Abs(p("Value") - 449# / 107) < 0.00001, "Estimated projection drifted"
 'A later independent observation cannot silently lose or double-count local history.
 b = RSAddBaseline(sid, "E6", Date + 1, 4, 200, 800, "New independent BOL observation")
 RSSetGroupBaseline g, sid, b, True
 Check RejectComplete(g), "Ambiguous manual reconciliation silently counted twice"
 trail = RSHistoryList(sid, "E6", summary)
 Check InStr(trail, "Earlier: Group completed: First") > 0 And InStr(summary, "No summary groups completed since") > 0 And RSBaselineSource(b) = "BOL " & Format$(Date + 1, "yyyy-mm-dd"), "A new BOL entry did not start a new history"
 'Model schema-27 XML on a disposable synthetic group. Migration must never post totals again.
 legacy = CreateSummaryGroup("Legacy accounting", "EVAL", "E4")
 legacyReport = Member(senior, "Legacy synthetic member"): AddGroupMember legacy, legacyReport, "E4"
 legacyBaseline = RSAddBaseline(sid, "E4", #12/31/2025#, 4, 7, 28, "Synthetic observation")
 RSSetGroupBaseline legacy, sid, legacyBaseline, True
 SetGroupWorkflow legacy, GroupStamp(legacy), "Completed"
 posting = RSPosting(legacy): b = RSLatestBaseline(sid, "E4")
 xml = "<SwiftGroupPolicy version=""1""><Workflow>Cancelled</Workflow><RSCAPostedBaseline>" & b & "</RSCAPostedBaseline><RSCAGroupSum>28</RSCAGroupSum><RSCAGroupCount>7</RSCAGroupCount><RSCAReportIDs>|" & legacyReport & "|</RSCAReportIDs><RSCACompletedAt>2026-10-22 01:00:00</RSCACompletedAt></SwiftGroupPolicy>"
 CurrentDb.Execute "UPDATE SummaryGroup SET KeySpec=" & QuoteSQL(xml) & " WHERE GroupID=" & legacy, dbFailOnError
 CurrentDb.Execute "DELETE FROM RSCountedReport WHERE CompletionID=" & posting, dbFailOnError
 CurrentDb.Execute "DELETE FROM GroupCompletionCurrent WHERE GroupID=" & legacy, dbFailOnError
 CurrentDb.Execute "DELETE FROM GroupCompletionReport WHERE CompletionID=" & posting, dbFailOnError
 CurrentDb.Execute "DELETE FROM GroupCompletion WHERE CompletionID=" & posting, dbFailOnError
 CurrentDb.Execute "DELETE FROM GroupLifecycle WHERE GroupID=" & legacy, dbFailOnError
 entries = DCount("*", "RSCABaseline")
 Call MigrateRSAccounting: MigrateRSAccounting
 posting = RSPosting(legacy)
 Check DCount("*", "RSCABaseline") = entries And DCount("*", "GroupCompletion", "GroupID=" & legacy) = 1, "Legacy migration double counted or duplicated"
 Check GroupWorkflow(legacy) = "Cancelled", "Legacy cancelled state was lost"
 Check IsNull(DLookup("RevisionID", "GroupCompletionReport", "CompletionID=" & posting)), "Legacy posting invented revision evidence"
 Check GroupPolicyValue(legacy, "RSCAGroupSum") = "28" And GroupPolicyValue(legacy, "RSCAGroupCount") = "7", "Legacy totals changed"
 Check DLookup("KeySpec", "SummaryGroup", "GroupID=" & legacy) = xml, "Legacy XML evidence was rewritten"
 Check Left$(RSReconcileAccounting(), 5) = "PASS:", "Mixed legacy/new accounting did not reconcile"
 TestRSCompletion = "PASS: " & n & " relational accounting, reconciliation, migration, idempotency and rollback assertions": Exit Function
Failed:
 TestRSCompletion = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
