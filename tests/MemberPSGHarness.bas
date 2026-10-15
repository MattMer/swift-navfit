Attribute VB_Name = "MemberPSGHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8370, , message
 n = n + 1
End Sub
Public Function TestMemberPSG(ByVal directory As String) As String
 Dim m As Long, other As Long, current As Long, id As Long, legacy As Long, placeholder As Long, groupID As Long, senior As Long, baseline As Long
 Dim p As Object, values As Object, previous As Object, result As Object, key As Variant, db As DAO.Database, a As DAO.Recordset, b As DAO.Recordset, field As DAO.Field, before As Long, stage As String, f As Access.Form
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC PDF MEMBER AND PSG", "Tester"
 m = SaveDirectoryMember(0, 0, "EXAMPLE, TEST", "EXAMPLE, TEST", "YN2", "000000007")
 Check DirectoryIdentityMatch("Different name", "Different name", "000000007") = m, "SSN match depends on name spelling"
 Check DirectoryIdentityMatch("EXAMPLE TEST", "", "") = m, "Unique normalized name not matched"
 Check DirectoryIdentityMatch("EXAMPLE TEST", "", "000000008") = 0, "Conflicting SSNs merged by name"
 other = SaveDirectoryMember(0, 0, "EXAMPLE, TEST", "EXAMPLE, TEST", "YN2", "000000008")
 Check DirectoryIdentityMatch("EXAMPLE TEST", "", "") = 0, "Ambiguous names merged"
 Check DirectoryIdentityMatch("EXAMPLE TEST", "", "000000007") = m, "SSN did not resolve same-name members"
 current = CommitDraft(0, "EVAL", "Current report", "EXAMPLE, TEST", #1/1/2027#, #9/30/2027#, "", administration:=Array("EXAMPLE, TEST", "YN2", Null, Null, Null), memberDetails:=Array("000000007", Null, Null, Null, Null, Null, Null))
 before = DCount("*", "CommandMember")
 Set p = PreviewStructuredPDF(directory & "\eval-expanded.pdf")
 id = ImportPDFPreview("20", "SSN")
 Check DCount("*", "CommandMember") = before, "PDF created placeholder directory record"
 Check MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & id) = m, "PDF linked to wrong member"
 Check MajorScalar("SELECT MatchState FROM MemberReport WHERE ReportID=" & id) = "Confirmed", "PDF identity unconfirmed"
 Set previous = PreviousMemberEvaluation(current)
 Check previous("ITA") <> "Unavailable" And previous("ITA") <> "Ambiguous", "Prior PDF not available for current report"
 Check previous("Promotion") = "Must Promote", "Prior PDF recommendation missing"
 Check MajorScalar("SELECT MemberSSN FROM CommandMember WHERE MemberID=" & m) = "000000007", "Existing member SSN lost"
 'Reproduce an old placeholder-linked import, then verify scoped repair.
 stage = "Create legacy placeholder": legacy = CommitDraft(0, "EVAL", "Legacy PDF", "Imported PDF member", Null, Null, "")
 placeholder = MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & legacy)
 Set values = CreateObject("Scripting.Dictionary"): values.Add "OfficialName", "EXAMPLE, TEST": values.Add "MemberSSN", "000000007": values.Add "RankRate", "YN2"
 stage = "Legacy final fields": Call CommitSavedFields(legacy, CurrentRevision(legacy), values, "Legacy import field pass")
 stage = "Legacy receipt": Set db = CurrentDb: Set a = db.OpenRecordset("SELECT * FROM ImportReceipt WHERE ReportID=" & id, dbOpenSnapshot): Set b = db.OpenRecordset("ImportReceipt", dbOpenDynaset)
 b.AddNew
 For Each field In a.Fields
  If (field.Attributes And dbAutoIncrField) = 0 Then b.Fields(field.Name) = field.Value
 Next
 b!ReceiptGuid = NewIdentity(): b!ReportID = legacy: b!SourceHash = SourceTextHash("Synthetic legacy PDF")
 b.Update: a.Close: b.Close
 stage = "Repair legacy": Check RepairPDFDirectoryLinks() = 1, "Legacy placeholder repair failed"
 Check MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & legacy) = m, "Legacy link not repaired"
 Check MajorScalar("SELECT IsArchived FROM CommandMember WHERE MemberID=" & placeholder), "Unused placeholder not retired"
 Check RepairPDFDirectoryLinks() = 0, "Repair not idempotent"
 Set values = RSSeniorValues(0): values("ReportingSeniorName") = "EXAMPLE, A B": values("ReportingSeniorSSN") = "000000009"
 senior = RSSaveSenior(0, "Synthetic senior", values)
 baseline = RSAddBaseline(senior, "E5", #10/1/2026#, 4, 70, 280, "Synthetic historical totals")
 groupID = CreateSummaryGroup("Synthetic E5", "EVAL", "E5"): AddGroupMember groupID, id, "E5"
 Check RSOpenGroupCalculator(groupID), "Open calculator"
 Check Not UXHasControl(Forms!frmUXGroupRSCA, "cmdSave"), "Redundant Save button remains"
 Forms!frmUXGroupRSCA!cboSenior = senior: RSGroupSeniorChanged
 Check MajorScalar("SELECT SeniorID FROM GroupRSCABaseline WHERE GroupID=" & groupID) = senior, "Senior selection not auto-saved"
 Forms!frmUXGroupRSCA!cboBaseline = baseline: RSGroupBaselineChanged
 Check Not Forms!frmUXGroupRSCA!chkExcludes, "New baseline retained old confirmation"
 Check MajorScalar("SELECT BaselineID FROM GroupRSCABaseline WHERE GroupID=" & groupID) = baseline, "Baseline not auto-saved"
 Forms!frmUXGroupRSCA!chkExcludes = True: RSSaveGroupCalculation
 Set result = RSGroupPSG(groupID)
 Check Not IsNull(result("Value")), "PSG unavailable after confirmed selection"
 Check InStr(UXGroupSeniorCaption(groupID), Format$(result("Value"), "0.00")) > 0, "Group caption missing PSG"
 Set values = CreateObject("Scripting.Dictionary"): values.Add "Trait33", "1.0"
 Call CommitSavedFields(id, CurrentRevision(id), values, "Changed trait")
 Set p = RSGroupPSG(groupID)
 Check CDbl(p("Value")) <> CDbl(result("Value")), "PSG did not recalculate saved traits"
 Check InStr(UXGroupSeniorCaption(groupID), Format$(p("Value"), "0.00")) > 0, "Caption PSG stale"
 stage = "Live group caption"
 DoCmd.OpenForm "frmUXComparison", , , , , , "G:" & groupID
 Set f = Forms!frmUXComparison!uxTraitGrid.Form
 f.Recordset.FindFirst "ReportID=" & id: f!Trait33 = "2.0"
 Check UXGridEdit("Trait33"), "Native trait autosave failed: " & Forms!frmUXComparison!lblStatus.Caption
 Set p = RSGroupPSG(groupID)
 Check InStr(Forms!frmUXComparison!uxSeniorCaption.Caption, "PSG RSCA: " & Format$(p("Value"), "0.00")) > 0, "Visible group PSG did not refresh after trait autosave"
 DoCmd.Close acForm, "frmUXComparison", acSaveNo
 Forms!frmUXGroupRSCA!cboBaseline = baseline: RSGroupBaselineChanged
 Set result = RSGroupPSG(groupID): Check IsNull(result("Value")), "Unconfirmed changed baseline calculated"
 DoCmd.Close acForm, "frmUXGroupRSCA", acSaveNo
 TestMemberPSG = "PASS: " & n & " PDF member identity/history and automatic PSG assertions": Exit Function
Failed:
 TestMemberPSG = "FAIL after " & n & " / " & stage & ": " & Err.Number & " " & Err.Description
End Function
Public Function TestRUADReportRefresh() As String
 Dim m As Long, id As Long, old As Long, values As Object, ruad As Object, version As Long, ssn As String
 On Error GoTo Failed
 Set ruad = RUADDecode(Null): ruad("PHONE") = "Synthetic retained phone": ruad("UIC") = "12345"
 m = SaveDirectoryMember(0, 0, "NAVY, JOSEPH G", "", "YNC", "", "RUAD", ruad)
 id = CommitDraft(0, "CHIEFEVAL", "Newer member report", "NAVY, JOSEPH G", #1/1/2027#, #9/30/2027#, "", administration:=Array("NAVY, JOSEPH G", "YNCS", Null, Null, Null), memberDetails:=Array("000000041", Null, Null, Null, Null, Null, Null))
 Check MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & id) = m, "RUAD name match created duplicate"
 Check MajorScalar("SELECT MemberSSN FROM CommandMember WHERE MemberID=" & m) = "000000041", "RUAD SSN not enriched"
 Check MajorScalar("SELECT RankRate FROM CommandMember WHERE MemberID=" & m) = "YNCS", "Newer report rank not adopted"
 Check MajorScalar("SELECT ReportName FROM CommandMember WHERE MemberID=" & m) = "NAVY, JOSEPH G", "Missing report name not enriched"
 Check MajorScalar("SELECT SourceKind FROM CommandMember WHERE MemberID=" & m) = "RUAD", "RUAD source overwritten"
 Set values = RUADDecode(MajorScalar("SELECT RUADData FROM CommandMember WHERE MemberID=" & m))
 Check values("PHONE") = "Synthetic retained phone" And values("UIC") = "12345", "Report enrichment erased RUAD details"
 old = CommitDraft(0, "CHIEFEVAL", "Older prior report", "NAVY, JOSEPH G", #1/1/2025#, #9/30/2025#, "", administration:=Array("NAVY, JOSEPH G", "YNC", Null, Null, Null), memberDetails:=Array("000000041", Null, Null, Null, Null, Null, Null))
 Check MajorScalar("SELECT RankRate FROM CommandMember WHERE MemberID=" & m) = "YNCS", "Older report downgraded current rank"
 version = MajorScalar("SELECT VersionCounter FROM CommandMember WHERE MemberID=" & m)
 RefreshDirectoryFromReport m, id
 Check MajorScalar("SELECT VersionCounter FROM CommandMember WHERE MemberID=" & m) = version, "No-op enrichment rewrote member"
 m = SaveDirectoryMember(0, 0, "SAMPLE, ALEX", "", "YN2", "", "RUAD")
 id = CommitDraft(0, "EVAL", "Name-only report", "SAMPLE, ALEX", #1/1/2027#, #9/30/2027#, "", administration:=Array("SAMPLE, ALEX", "YN1", Null, Null, Null))
 Check MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & id) = m, "Name-only RUAD identity not linked"
 Set values = CreateObject("Scripting.Dictionary"): values.Add "MemberSSN", "000000042"
 Call CommitSavedFields(id, CurrentRevision(id), values, "Add member SSN later")
 Check MajorScalar("SELECT MemberSSN FROM CommandMember WHERE MemberID=" & m) = "000000042", "Later SSN not added to RUAD profile"
 Check MajorScalar("SELECT MatchState FROM MemberReport WHERE ReportID=" & id) = "Confirmed", "Later SSN made established name link unresolved"
 TestRUADReportRefresh = "PASS: " & n & " combined member/PSG and RUAD enrichment assertions": Exit Function
Failed:
 TestRUADReportRefresh = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
