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
Private Sub Counts(ByVal senior As Long, ByVal count As Long, ByVal total As Double)
 Dim id As Long
 id = RSLatestBaseline(senior, "E6")
 Check CLng(DLookup("TraitCount", "RSCABaseline", "BaselineID=" & id)) = count, "Wrong running trait count"
 Check Abs(CDbl(DLookup("TraitSum", "RSCABaseline", "BaselineID=" & id)) - total) < 0.00001, "Wrong running trait sum"
End Sub
Public Function TestRSCompletion() As String
 Dim senior As Object, sid As Long, b As Long, g As Long, h As Long, duplicate As Long, a As Long, c As Long, key As Variant, f As Object, p As Object, control As Access.Control, entries As Long, rev As Long, nob As Long, estimated As Long, late As Long, original As String
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
 rev = CurrentRevision(a)
 SetGroupWorkflow g, GroupStamp(g), "Completed": Counts sid, 107, 448
 Check GroupWorkflow(g) = "Completed" And CurrentRevision(a) = rev, "Completion changed report revision"
 Check DLookup("TraitCount", "RSCABaseline", "BaselineID=" & b) = 100, "Original baseline overwritten"
 entries = DCount("*", "RSCABaseline")
 SetGroupWorkflow g, GroupStamp(g), "Completed"
 Check DCount("*", "RSCABaseline") = entries, "Repeat completion double counted"
 Set p = RSGroupPSG(g): Check Abs(p("Value") - 448# / 107) < 0.00001, "Completed group counted twice in projection"
 h = CreateSummaryGroup("Second", "EVAL", "E6"): AddGroupMember h, c, "E6": RSSetGroupBaseline h, sid, b, True
 Set p = RSGroupPSG(h): Check Abs(p("Value") - 476# / 114) < 0.00001, "Projection omitted earlier completed group"
 SetGroupWorkflow h, GroupStamp(h), "Completed": Counts sid, 114, 476
 SetGroupWorkflow g, GroupStamp(g), "In-Work": Counts sid, 114, 476
 Set f = CreateObject("Scripting.Dictionary")
 For Each key In TraitKeys(): f.Add CStr(key), "3.0": Next
 CommitSavedFields a, CurrentRevision(a), f, "Correct completed traits"
 SetGroupWorkflow g, GroupStamp(g), "Completed": Counts sid, 114, 469
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
 'Notes and Review must display different content in the actual editor.
 DoCmd.OpenForm "frmEditor", , , , , , CStr(a)
 Check UXSelectSection(7), "Notes selection failed"
 Check Forms!frmEditor!uxFieldSurface.Visible, "Draft notes hidden"
 Check UXSelectSection(8), "Review selection failed"
 Check Not Forms!frmEditor!uxFieldSurface.Visible, "Review retains notes surface"
 For Each control In Forms!frmEditor!uxFieldSurface.Form.Controls
  If control.Name <> "uxFocusAnchor" Then Check Not control.Visible, "Review retains child content: " & control.Name
 Next
 Check Forms!frmEditor!cmdRoute.Visible And Forms!frmEditor!cmdComments.Visible And Forms!frmEditor!cmdChecklist.Visible, "Review actions absent"
 Check UXSelectSection(9), "History selection failed"
 Check Not Forms!frmEditor!uxFieldSurface.Visible And Forms!frmEditor!lstRevisions.Visible And Not Forms!frmEditor!cmdRoute.Visible, "History visibility incorrect"
 Check UXSelectSection(0), "Front selection failed"
 Check Forms!frmEditor!uxFieldSurface.Visible And Not Forms!frmEditor!lstRevisions.Visible, "Front retains history"
 Check UXSelectSection(7) And UXSelectSection(8), "Repeat tab transition failed"
 Check Not Forms!frmEditor!uxFieldSurface.Visible, "Repeat Review retains notes"
 TestRSCompletion = "PASS: " & n & " completion accounting and tab visibility assertions": Exit Function
Failed:
 TestRSCompletion = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
