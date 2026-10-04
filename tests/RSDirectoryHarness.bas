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
Public Function TestRSDirectory() As String
    Dim senior As Long, values As Object, saved As Object, reportID As Long, changes As Object, groupID As Long, baseline As Long, result As Object, key As Variant, operation As Long, f As Access.Form, stage As String, snapshot As Long
    On Error GoTo Failed
    checks = 0: stage = "configure": ConfigureWorkspace "SYNTHETIC RS DIRECTORY", "Synthetic Tester"
    Set values = RSSeniorValues(0): values("ReportingSeniorName") = "O'NEIL A": values("ReportingSeniorSSN") = "000123456": values("ReportingSeniorAddress1") = "123 SYNTHETIC ST": values("ReportingSeniorLastName") = "OTHER"
    stage = "save": senior = RSSaveSenior(0, "Synthetic O'Neil", values)
    Set saved = RSSeniorValues(senior)
    Check saved("ReportingSeniorName") = "O'NEIL A" And saved("ReportingSeniorLastName") = "OTHER", "Independent full name and name parts changed"
    Check DCount("*", "ReportingSeniorValue", "SeniorID=" & senior) = 16, "Directory field coverage incorrect"
    values("ReportingSeniorTitle") = "COMMANDING OFF": Call RSSaveSenior(senior, "Synthetic RS", values, 1)
    Check RejectSave(senior, values, 1), "Stale record update accepted"
    Check MajorScalar("SELECT VersionCounter FROM ReportingSeniorRecord WHERE SeniorID=" & senior) = 2, "Stale update changed version"
    CurrentDb.TableDefs("ReportingSeniorValue").Fields("FieldKey").ValidationRule = "<> 'ReportingSeniorTitle'"
    Check RejectSave(senior, values, 2), "Injected directory failure accepted"
    Check MajorScalar("SELECT VersionCounter FROM ReportingSeniorRecord WHERE SeniorID=" & senior) = 2, "Failed save did not roll back version"
    Check DCount("*", "ReportingSeniorValue", "SeniorID=" & senior) = 16, "Failed save lost directory values"
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
    Call RSSelectSeniorPage(2): Check f!rsReportingSeniorLastName.Visible, "Name parts section hidden"
    Call RSSelectSeniorPage(3): f!cboPaygrade = "E6": f!txtAsOf = "2026-02-30": f!txtRSCA = "4.00": f!txtTraitCount = "10"
    Check Not RSAddBaselineUI(), "Invalid baseline date accepted"
    Check f!txtAsOf = "2026-02-30", "Baseline failure lost pending values"
    f!txtAsOf = "2026-04-01": Check RSAddBaselineUI(), "Baseline UI save failed"
    DoCmd.Close acForm, "frmUXSeniorEditor", acSaveNo
    stage = "comparison UI": DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = groupID
    snapshot = DCount("*", "GroupSnapshot"): Check UXOpenComparison(), "Automatic comparison failed"
    Check DCount("*", "GroupSnapshot") = snapshot + 1, "Comparison did not automatically capture baseline"
    Forms!frmUXComparison!lstTraits = reportID
    Check Forms!frmUXComparison!lstTraits.Column(0) = reportID, "Comparison roster not loaded"
    Check UXSelectComparisonMember(), "Inline comparison controls did not load"
    stage = "group calculator UI"
    Check RSOpenGroupCalculator(), "Group calculator failed to open"
    Check Forms!frmUXGroupRSCA!cboSenior = senior And Forms!frmUXGroupRSCA!cboBaseline = baseline, "Saved baseline selection not restored"
    Check RSStageGroupSenior(), "Group senior details failed to stage"
    Check UXPreviewBulk(), "Group senior detail preview failed"
    Check UXApplyBulk(), "Group senior detail atomic apply failed"
    Set saved = SavedFields(reportID)
    Check saved("ReportingSeniorSSN") = "000123456" And saved("ReportingSeniorName") = "O'NEIL A", "Group apply lost senior details"
    Check RSShowGroupCalculation(), "Group PSG display failed"
    Check InStr(Forms!frmUXGroupRSCA!lblCalculation.Caption, "3.00") > 0, "Group PSG display not recalculated"
    TestRSDirectory = "PASS: " & checks & " reporting senior, RSCA and comparison assertions": Exit Function
Failed:
    TestRSDirectory = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
