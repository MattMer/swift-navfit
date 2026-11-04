Attribute VB_Name = "RSDirectoryHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 4900, , message
    checks = checks + 1
End Sub
Private Function RejectSave(ByVal id As Long, ByVal values As Object, ByVal version As Long) As Boolean
    On Error GoTo Rejected
    Call RSSaveSenior(id, "Rejected", values, version)
    Exit Function
Rejected:
    RejectSave = True
End Function
Private Function RejectBaseline(ByVal id As Long, ByVal grade As String, ByVal average As Variant, ByVal count As Long, ByVal total As Variant) As Boolean
    On Error GoTo Rejected
    Call RSAddBaseline(id, grade, #1/1/2026#, average, count, total)
    Exit Function
Rejected:
    RejectBaseline = True
End Function
Private Function RejectLink(ByVal member As Long) As Boolean
    On Error GoTo Rejected
    Call RSSaveSenior(0, "Rejected link", RSSeniorValues(0), 0, member)
    Exit Function
Rejected:
    RejectLink = True
End Function
'A senior linked to a command member: mapping, picker, link status, update, unlink,
'one senior per member, and the list columns.
Private Sub TestMemberLink()
    Dim member As Long, archived As Long, ruad As Object, mapped As Object, f As Access.Form, linked As Long, i As Long, row As Long
    Set ruad = RUADDecode(Null): ruad("DESG") = "1110": ruad("UIC") = "12345"
    member = SaveDirectoryMember(0, 0, "Sam Rivera", "RIVERA, SAM JOSE", "CDR", "000000777", "Manual", ruad)
    Set mapped = RSMemberValues(member)
    Check mapped("ReportingSeniorName") = "RIVERA, S J" And mapped("ReportingSeniorLastName") = "RIVERA" And mapped("ReportingSeniorFirstInitial") = "S" And mapped("ReportingSeniorMiddleInitial") = "J", "Member name not mapped to the senior block"
    Check mapped("ReportingSeniorGrade") = "CDR" And mapped("ReportingSeniorDesignator") = "1110" And mapped("ReportingSeniorUIC") = "12345" And mapped("ReportingSeniorSSN") = "000000777", "Member details not mapped"
    Check Not mapped.Exists("ReportingSeniorTitle"), "Profile supplied a billet title"
    DoCmd.OpenForm "frmUXSeniors"
    Check RSOpenSenior(True) And CurrentProject.AllForms("frmUXSeniorPicker").IsLoaded, "New reporting senior did not open the member picker"
    Forms!frmUXSeniorPicker!txtSearch = "12345": Check RSPickerFilter(), "Picker search failed"
    Check Forms!frmUXSeniorPicker!lstMembers.ListCount = 1 And Forms!frmUXSeniorPicker!lstMembers.Column(3, 0) = "1110" And Forms!frmUXSeniorPicker!lstMembers.Column(4, 0) = "12345", "Picker did not find the member by UIC with designator and UIC"
    Check Not RSPickerUse() And CurrentProject.AllForms("frmUXSeniorPicker").IsLoaded, "Picker used no selection"
    Forms!frmUXSeniorPicker!lstMembers = member: Check RSPickerUse(), "Use selected member failed"
    Set f = Forms!frmUXSeniorEditor
    Check f!rsReportingSeniorUIC = "12345" And f!rsReportingSeniorGrade = "CDR" And f!txtMemberID = member And RSSeniorDirty(), "Picker did not fill a pending linked record"
    Check InStr(f!lblDirectoryLink.Caption, "Matches the directory") > 0, "Link status missing"
    f!rsReportingSeniorTitle = "CO": f!cmdSave.SetFocus: Check RSSaveSeniorUI(), "Linked senior save failed"
    linked = f!txtSeniorID
    Check DLookup("MemberID", "ReportingSeniorRecord", "SeniorID=" & linked) = member And Not RSSeniorDirty(), "Directory link not saved"
    CurrentDb.Execute "UPDATE CommandMember SET RankRate='CAPT' WHERE MemberID=" & member, dbFailOnError
    Check RSShowLink() And InStr(f!lblDirectoryLink.Caption, "Directory differs: Grade") > 0, "Directory difference not shown"
    Check RSLinkUpdate() And f!rsReportingSeniorGrade = "CAPT" And f!rsReportingSeniorTitle = "CO" And RSSeniorDirty(), "Update from directory changed the wrong values"
    Check RSUnlinkMember() And IsNull(f!txtMemberID) And InStr(f!lblDirectoryLink.Caption, "Not linked") > 0, "Unlink not pending"
    DoCmd.Close acForm, "frmUXSeniorEditor", acSaveNo
    Check DLookup("MemberID", "ReportingSeniorRecord", "SeniorID=" & linked) = member, "Closing without saving changed the link"
    Check RejectLink(member), "Second senior record linked to the same member"
    archived = SaveDirectoryMember(0, 0, "Pat Archived", "ARCHIVED, PAT", "LT", "")
    CurrentDb.Execute "UPDATE CommandMember SET IsArchived=True WHERE MemberID=" & archived, dbFailOnError
    Check RejectLink(archived), "Archived member linked"
    Check RSOpenSenior(True), "Picker reopen failed"
    Forms!frmUXSeniorPicker!txtSearch = "RIVERA": RSPickerFilter
    Check Forms!frmUXSeniorPicker!lstMembers.Column(5, 0) = "Already a senior", "Picker did not mark an existing senior"
    Forms!frmUXSeniorPicker!lstMembers = member
    Check Not RSPickerUse() And InStr(Forms!frmUXSeniorPicker!lblStatus.Caption, "already the reporting senior") > 0, "Picker created a duplicate senior"
    archived = SaveDirectoryMember(0, 0, "Quinn Blankname", "BLANKNAME, QUINN", "LT", "")
    CurrentDb.Execute "UPDATE CommandMember SET ReportName='' WHERE MemberID=" & archived, dbFailOnError
    Forms!frmUXSeniorPicker!txtSearch = "Blankname": RSPickerFilter
    Check Forms!frmUXSeniorPicker!lstMembers.Column(1, 0) = "Quinn Blankname", "Picker showed a blank name for an empty report name"
    Check RSPickerManual() And Forms!frmUXSeniorEditor!txtSeniorID = 0 And IsNull(Forms!frmUXSeniorEditor!txtMemberID), "Enter details by hand did not open a blank record"
    DoCmd.Close acForm, "frmUXSeniorEditor", acSaveNo
    Check RSLoadDirectory(), "Directory list failed"
    With Forms!frmUXSeniors!lstSeniors
        row = -1
        For i = 0 To .ListCount - 1
            If Val(.Column(0, i)) = linked Then row = i
        Next
        Check row >= 0, "Linked senior missing from the list"
        Check .Column(2, row) = "CDR" And .Column(3, row) = "1110" And .Column(4, row) = "CO" And .Column(5, row) = "12345" And .Column(6, row) = "Linked", "List columns wrong"
    End With
    Check UXSortList("frmUXSeniors", "lstSeniors", 5), "Senior list did not sort by UIC"
    DoCmd.Close acForm, "frmUXSeniors", acSaveNo
End Sub
Public Function TestRSDirectory() As String
    Dim senior As Long, values As Object, saved As Object, reportID As Long, changes As Object, groupID As Long, baseline As Long, result As Object, key As Variant, operation As Long, f As Access.Form, stage As String, snapshot As Long
    On Error GoTo Failed
    checks = 0: stage = "configure": ConfigureWorkspace "SYNTHETIC RS DIRECTORY", "Synthetic Tester"
    Set values = RSSeniorValues(0): values("ReportingSeniorName") = "O'NEIL A": values("ReportingSeniorSSN") = "000123456": values("ReportingSeniorAddress1") = "123 SYNTHETIC ST": values("ReportingSeniorLastName") = "OTHER"
    stage = "save": senior = RSSaveSenior(0, "Synthetic O'Neil", values)
    Set saved = RSSeniorValues(senior)
    Check saved("ReportingSeniorName") = "O'NEIL A" And saved("ReportingSeniorLastName") = "OTHER", "Independent full name and name parts changed"
    Check DCount("*", "ReportingSeniorValue", "SeniorID=" & senior) = 17, "Directory field coverage incorrect"
    values("ReportingSeniorTitle") = "COMMANDING OFF": Call RSSaveSenior(senior, "Synthetic RS", values, 1)
    Check RejectSave(senior, values, 1), "Stale record update accepted"
    Check MajorScalar("SELECT VersionCounter FROM ReportingSeniorRecord WHERE SeniorID=" & senior) = 2, "Stale update changed version"
    CurrentDb.TableDefs("ReportingSeniorValue").Fields("FieldKey").ValidationRule = "<> 'ReportingSeniorTitle'"
    Check RejectSave(senior, values, 2), "Injected directory failure accepted"
    Check MajorScalar("SELECT VersionCounter FROM ReportingSeniorRecord WHERE SeniorID=" & senior) = 2, "Failed save did not roll back version"
    Check DCount("*", "ReportingSeniorValue", "SeniorID=" & senior) = 17, "Failed save lost directory values"
    CurrentDb.TableDefs("ReportingSeniorValue").Fields("FieldKey").ValidationRule = ""
    Check RejectBaseline(senior, "E5", 4.79, 124, 593), "Mismatched exact totals accepted"
    Check RejectBaseline(senior, "E5", 6, 124, Null), "Out-of-range RSCA accepted"
    Check RejectBaseline(senior, "E5", 4, -1, Null), "Negative trait count accepted"
    stage = "baseline": baseline = RSAddBaseline(senior, "E5", #1/1/2026#, 4.79, 124, 594, "Synthetic exact BOL totals")
    reportID = CommitDraft(0, "EVAL", "Synthetic report", "Synthetic member", Null, Null, "", "Test")
    Set changes = CreateObject("Scripting.Dictionary")
    For Each key In TraitKeys(): changes.Add CStr(key), "3.0": Next
    changes.Add "ObservationStatus", "Observed": changes.Add "ReportingSeniorSSN", "000123456"
    Call CommitSavedFields(reportID, CurrentRevision(reportID), changes, "Synthetic traits")
    groupID = CreateSummaryGroup("Synthetic E5", "EVAL", "E5"): AddGroupMember groupID, reportID, "E5"
    RSSetGroupBaseline groupID, senior, baseline, False
    Set result = RSGroupPSG(groupID): Check IsNull(result("Value")), "Unconfirmed exclusion calculated"
    RSSetGroupBaseline groupID, senior, baseline, True
    Set result = RSGroupPSG(groupID): Check Abs(result("Value") - 615# / 131) < 0.0000001, "Incorrect PSG weighted totals"
    Check Not result("Estimated"), "Exact historical totals labelled estimated"
    Check DLookup("TraitCount", "RSCABaseline", "BaselineID=" & baseline) = 124, "Calculation advanced historical count"
    baseline = RSAddBaseline(senior, "E5", #2/1/2026#, 4.79, 124, Null)
    RSSetGroupBaseline groupID, senior, baseline, True
    Set result = RSGroupPSG(groupID): Check result("Estimated"), "Rounded average calculation not labelled estimate"
    baseline = RSAddBaseline(senior, "E5", #3/1/2026#, Null, 0, 0)
    RSSetGroupBaseline groupID, senior, baseline, True
    Set result = RSGroupPSG(groupID): Check result("Value") = 3, "First group with zero prior traits incorrect"
    changes.RemoveAll: changes.Add "ObservationStatus", "NOB": Call CommitSavedFields(reportID, CurrentRevision(reportID), changes, "NOB")
    Set result = RSGroupPSG(groupID): Check IsNull(result("Value")), "Whole-report NOB contributed traits"
    changes("ObservationStatus") = "Observed": changes.Add "ReportingSeniorSSN", "111111111": Call CommitSavedFields(reportID, CurrentRevision(reportID), changes, "Different RS")
    Set result = RSGroupPSG(groupID): Check IsNull(result("Value")), "Different reporting senior calculated"
    stage = "directory UI": DoCmd.OpenForm "frmUXSeniors": Check RSLoadDirectory(), "Directory list failed"
    Forms!frmUXSeniors!lstSeniors = senior: Check RSOpenSenior(), "Directory record failed to open"
    Set f = Forms!frmUXSeniorEditor: Check f!rsReportingSeniorSSN = "000123456", "Directory UI lost SSN"
    Call RSSelectSeniorPage(1): Check f!rsReportingSeniorAddress1.Visible, "Address section hidden"
    Call RSSelectSeniorPage(2): Check Not f!rsReportingSeniorLastName.Visible, "Redundant name parts exposed"
    Call RSSelectSeniorPage(3): f!cboPaygrade = "E6": f!txtAsOf = "2026-02-30": f!txtRSCA = "4.00": f!txtTraitCount = "10"
    Check Not RSAddBaselineUI(), "Invalid baseline date accepted"
    Check f!txtAsOf = "2026-02-30", "Baseline failure lost pending values"
    f!txtAsOf = "2026-04-01": Check RSAddBaselineUI(), "Baseline UI save failed"
    DoCmd.Close acForm, "frmUXSeniorEditor", acSaveNo
    stage = "directory link": TestMemberLink
    stage = "comparison UI": DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = groupID
    Forms!frmMajorWorkbench!uxMode = "Comparison"
    snapshot = DCount("*", "GroupSnapshot"): Check UXOpenComparison(), "Automatic comparison failed"
    Check DCount("*", "GroupSnapshot") = snapshot + 1, "Comparison did not automatically capture baseline"
    Forms!frmUXComparison!uxTraitGrid.Form.Recordset.FindFirst "ReportID=" & reportID
    Check Forms!frmUXComparison!uxTraitGrid.Form!ReportID = reportID, "Comparison roster not loaded"
    Check UXSelectComparisonMember(), "Inline comparison controls did not load"
    stage = "group calculator UI"
    Check RSOpenGroupCalculator(), "Group calculator failed to open"
    Check Forms!frmUXGroupRSCA!cboSenior = senior And Forms!frmUXGroupRSCA!cboBaseline = baseline, "Saved baseline selection not restored"
    Check InStr(Forms!frmUXGroupRSCA!lblStatus.Caption, "selected reporting senior SSN") > 0, "Initial unavailable reason missing"
    Check InStr(Forms!frmUXGroupRSCA!txtCalculationContext, "Manual historical baseline") > 0, "Historical baseline source missing"
    Check RSStageGroupSenior(vbYes), "Group senior details failed to apply"
    Check InStr(Forms!frmUXGroupRSCA!lblStatus.Caption, "saved") > 0, "Group senior result missing"
    Set saved = SavedFields(reportID)
    Check saved("ReportingSeniorSSN") = "000123456" And saved("ReportingSeniorName") = "O'NEIL A", "Group apply lost senior details"
    Check RSShowGroupCalculation(), "Group PSG display failed"
    Check InStr(Forms!frmUXGroupRSCA!lblCalculation.Caption, "3.00") > 0, "Group PSG display not recalculated"
    ' The report's "Use record in pending draft" reaches the visible page and the save.
    stage = "editor apply"
    Dim drafted As Long
    drafted = CommitDraft(0, "EVAL", "Use senior", "SYNTHETIC USE", #1/1/2026#, #9/30/2026#, "")
    DoCmd.OpenForm "frmEditor", , , , , , CStr(drafted)
    Forms!frmEditor!uxSeniorRecord = senior
    Check RSUseSeniorInEditor(), "Use record in pending draft failed"
    Check Forms!frmEditor!uxFieldSurface.Form!txtReportingSeniorName.Value = "O'NEIL A", "The visible field did not show the senior record"
    Check UXSaveEditor(), "Save after using the senior record failed"
    Set saved = SavedFields(drafted)
    Check saved("ReportingSeniorSSN") = "000123456" And saved("ReportingSeniorAddress1") = "123 SYNTHETIC ST", "The senior record did not reach the saved report"
    Forms!frmEditor.OnUnload = "": DoCmd.Close acForm, "frmEditor", acSaveNo
    TestRSDirectory = "PASS: " & checks & " reporting senior, RSCA and comparison assertions": Exit Function
Failed:
    TestRSDirectory = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
