Attribute VB_Name = "ReadinessHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8500, , message
 n = n + 1
End Sub
Private Function Has(ByVal findings As Collection, ByVal code As String, Optional ByVal field As String = "") As Boolean
 Dim item As Object
 For Each item In findings
  If item("Code") = code And (field = "" Or item("Key") = field) Then Has = True: Exit Function
 Next
End Function
Private Function SavedRejected(ByVal id As Long, ByVal pack As Long) As Boolean
 On Error GoTo Rejected
 RunRulePack id, pack: Exit Function
Rejected:
 SavedRejected = True
End Function
Private Sub CheckApplicability(ByVal id As Long)
 Dim row As Variant, f As Object, result As Collection, item As Object, expected As Date
 Dim cases As String, parts As Variant
 ' Independently transcribed from CH-1 tables 1 and 1A in the retained compilation.
 cases = "O1:5:31|O2:2:28|O3:1:31|O4:10:31|O5:4:30|O6:7:31|W1:9:30|W2:9:30|W3:3:31|W4:3:31|W5:3:31|E1:7:15|E2:7:15|E3:7:15|E4:6:15|E5:3:15|E6:11:15|E7:9:15|E8:9:15|E9:4:15"
 For Each row In Split(cases, "|")
  parts = Split(CStr(row), ":"): expected = DateSerial(2026, CLng(parts(1)), CLng(parts(2)))
  Check PeriodicExpectedDate(CStr(parts(0)), #12/31/2026#) = expected, "Table 1: " & parts(0)
  expected = DateSerial(2027, CLng(parts(1)), CLng(parts(2)))
  If parts(0) = "O5" Then expected = #6/30/2027#
  If parts(0) = "O6" Then expected = #3/31/2027#
  Check PeriodicExpectedDate(CStr(parts(0)), #1/1/2027#) = expected, "Table 1A: " & parts(0)
 Next
 Check PeriodicScheduleVersion(#12/15/2025#) = "", "Unverified historical calendar claimed"
 Check PeriodicScheduleVersion(#12/16/2025#) = "CH1-T1", "CH-1 start boundary"
 Check PeriodicExpectedDate("O5", #12/31/2025#) = #4/30/2025#, "Pre-2026 calendar selection"
 Check IsNull(PeriodicExpectedDate("E8", #1/1/2024#, "CMDCS")), "CMDCS bypassed coverage"
 Check IsNull(PeriodicExpectedDate("O7", #1/1/2027#)), "Unknown grade calendar fabricated"
 Check IsNull(PeriodicExpectedDate("E7", #1/1/2027#, "CMDCS")), "CMDCS wrong grade accepted"
 Check PeriodicExpectedDate("O2", #2/28/2027#) = #2/28/2027#, "Nonleap February"
 Check PeriodicExpectedDate("O2", #2/29/2028#) = #2/29/2028#, "Leap February"
 Check PeriodicExpectedDate("O2", #2/28/2100#) = #2/28/2100#, "Nonleap century February"
 For Each row In Array(Array("CHIEFEVAL", "ITCM", ""), Array("FITREP", "CWO5", "7310"), Array("FITREP", "CAPT", "6110"))
  Set f = Example(id, CStr(row(0))): f("RankRate") = row(1): f("Designator") = row(2): f("OccasionPeriodic") = "Yes"
  f("PeriodEnd") = #5/26/2026#: Set result = ValidateLiveFields(f)
  Check Not Has(result, "VR-PERIODIC", "OccasionPeriodic"), "Exemption predates NAVADMIN: " & row(1)
  f("PeriodEnd") = #6/2/2026#: Set result = ValidateLiveFields(f)
  Check Not Has(result, "VR-PERIODIC", "OccasionPeriodic"), "Exemption before boundary: " & row(1)
  f("PeriodEnd") = #6/3/2026#: Set result = ValidateLiveFields(f)
  Check Has(result, "VR-PERIODIC", "OccasionPeriodic"), "Exemption missing at boundary: " & row(1)
  For Each item In result
   If item("Code") = "VR-PERIODIC" Then Check InStr(item("Reference"), "NAVADMIN 132/26") > 0 And InStr(item("Message"), "3 June 2026") > 0, "Exemption source/date missing"
  Next
  f("ReportRegular") = "No": f("ReportConcurrent") = "Yes": Set result = ValidateLiveFields(f)
  Check Not Has(result, "VR-PERIODIC", "OccasionPeriodic"), "Regular exemption applied to pure Concurrent"
 Next
 Set f = Example(id, "EVAL"): f("PeriodEnd") = #12/31/2024#: Set result = ValidateLiveFields(f)
 Check Has(result, "VR-APPLICABILITY"), "Unsupported historical report silently passed"
 For Each item In result
  If item("Code") = "VR-APPLICABILITY" Then Check item("Severity") = "Needs review", "Coverage gap treated as invalid report"
 Next
 f("PeriodEnd") = #5/25/2026#: Set result = ValidateLiveFields(f)
 Check Has(result, "VR-APPLICABILITY"), "General policy start boundary missing"
 f("PeriodEnd") = #5/26/2026#: Set result = ValidateLiveFields(f)
 Check Not Has(result, "VR-APPLICABILITY"), "General policy start rejected"
 f("PeriodEnd") = Null: Set result = ValidateLiveFields(f)
 Check Has(result, "VR-APPLICABILITY"), "Missing date coverage not disclosed"
 Set f = Example(id, "FITREP"): f("RankRate") = "CDR": f("OccasionPeriodic") = "Yes": f("PeriodEnd") = #6/30/2027#
 Set result = ValidateLiveFields(f): Check Not Has(result, "VR-PERIODIC"), "Future O5 calendar rejected"
 f("PeriodEnd") = #4/30/2027#: Set result = ValidateLiveFields(f)
 Check Has(result, "VR-PERIODIC"), "Obsolete O5 date not flagged"
End Sub
Private Sub CheckHistoricalPack(ByVal id As Long)
 Dim db As DAO.Database, r As DAO.Recordset, oldPack As Long, oldRun As Long, before As Long, currentPack As Long
 Set db = CurrentDb: Set r = db.OpenRecordset("RulePack", dbOpenDynaset): r.AddNew
 r!VersionName = "1610.10H CH-2 / shared-readiness-r110": r!SourceTitle = "Historical sentinel"
 r!SourceSHA256 = NAVY_RULE_SOURCE_HASH: r!EffectiveFrom = #5/26/2026#: r!EffectiveTo = Null
 r.Update: r.Bookmark = r.LastModified: oldPack = r!PackID: r.Close
 Set r = db.OpenRecordset("ValidationRun", dbOpenDynaset): r.AddNew
 r!ReportID = id: r!RevisionID = CurrentRevision(id): r!PackID = oldPack: r!RunAt = #10/1/2026#
 r.Update: r.Bookmark = r.LastModified: oldRun = r!ValidationRunID: r.Close
 Set r = db.OpenRecordset("ValidationFinding", dbOpenDynaset): r.AddNew
 r!ValidationRunID = oldRun: r!FindingState = "NeedsReview": r!Message = "Historical finding sentinel": r.Update: r.Close
 before = DCount("*", "ValidationRun"): currentPack = InstallNavyDraftChecks()
 Check SavedRejected(id, oldPack), "Old engine rerun silently used new engine"
 Check DCount("*", "ValidationRun") = before, "Rejected old rerun wrote history"
 Check currentPack <> oldPack And InstallNavyDraftChecks() = currentPack, "Pack version/idempotency"
 Check MajorScalar("SELECT SourceTitle FROM RulePack WHERE PackID=" & oldPack) = "Historical sentinel", "Historical pack changed"
 Check MajorScalar("SELECT EffectiveFrom FROM RulePack WHERE PackID=" & oldPack) = #5/26/2026#, "Historical effective date changed"
 Check MajorScalar("SELECT RunAt FROM ValidationRun WHERE ValidationRunID=" & oldRun) = #10/1/2026#, "Historical run changed"
 Check MajorScalar("SELECT Message FROM ValidationFinding WHERE ValidationRunID=" & oldRun) = "Historical finding sentinel", "Historical result changed"
End Sub
Private Function Example(ByVal id As Long, ByVal kind As String) As Object
 Dim f As Object, key As Variant
 Set f = SavedFields(id)
 f("FormKind") = kind: f("OfficialName") = "EXAMPLE, A": f("RankRate") = IIf(kind = "EVAL", "IT1", IIf(kind = "CHIEFEVAL", "ITC", "LT"))
 f("MemberSSN") = "000000001": f("ReportingSeniorSSN") = "000000002": f("ReportingSeniorName") = "SENIOR, A"
 f("ReportingSeniorGrade") = "CDR": f("ReportingSeniorDesignator") = "1310": f("ReportingSeniorTitle") = "CO"
 f("ReportingSeniorUIC") = "12345": f("UnitUIC") = "12345": f("ShipStation") = "SYNTHETIC COMMAND"
 f("PeriodStart") = #1/1/2026#: f("PeriodEnd") = #9/30/2026#: f("DateReported") = "2026-01-01": f("PromotionStatus") = "REGULAR"
 f("ComponentActive") = "Yes": f("ReportRegular") = "Yes": f("OccasionDetachment") = "Yes": f("OccasionPeriodic") = "No"
 f("ObservationStatus") = "Observed": f("PromotionRecommendation") = "Promotable": f("Designator") = IIf(kind = "FITREP", "1110", "AW")
 For Each key In TraitKeys(): f(CStr(key)) = "4.0": Next
 f("AssignmentRecommendationA") = "TECHNICIAN": f("AssignmentRecommendationB") = Null
 If kind = "CHIEFEVAL" Then f("AssignmentRecommendationA") = Null
 f("RetentionYes") = "Yes": f("RetentionNo") = "No": f("OfficialComments") = "Synthetic comments."
 f("CounselingDate") = "2026-06-01": f("CounselorName") = "COUNSELOR, A": f("ReadinessSummary") = "P": f("ReadinessDate") = "2026-06-01"
 f("PrimaryDuty") = "TECHNICIAN": f("DutyDescription") = "Synthetic duty description.": f("CommandEmployment") = "Synthetic command employment."
 f("ReportingSeniorAddress1") = "123 TEST ST": f("ReportingSeniorCity") = "TEST": f("ReportingSeniorState") = "CA": f("ReportingSeniorZip") = "12345": f("ReportingSeniorPhone") = "0000000000"
 f("RaterName") = "RATER, A, ITC": f("SeniorRaterName") = "NONE AVAILABLE"
 Set Example = f
End Function
Private Sub CheckReadinessPaths(ByVal id As Long, ByVal groupID As Long)
 Dim changes As Object, original As Object, pending As Object, overrides As Object, result As Collection, item As Object, preview As Object
 Dim pack As Long, run As Long, revisions As Long, coverage As Long, exportError As Long, ignored As String
 Set original = SavedFields(id): Set changes = CreateObject("Scripting.Dictionary")
 changes("PeriodStart") = #1/1/2024#: changes("PeriodEnd") = #12/31/2024#
 CommitSavedFields id, CurrentRevision(id), changes, "Synthetic historical policy coverage"
 pack = InstallNavyDraftChecks(): run = RunRulePack(id, pack)
 Check DCount("*", "ValidationFinding", "ValidationRunID=" & run & " AND FindingState='NeedsReview' AND Message Like '*VR-APPLICABILITY*'") = 1, "Saved coverage finding missing"
 Check DCount("*", "ValidationFinding", "ValidationRunID=" & run & " AND FindingState='Pass'") = 0, "Unsupported saved run claimed pass"
 Set result = ValidateLiveGroup(groupID): coverage = 0
 For Each item In result
  If item("ReportID") = id And item("Code") = "VR-APPLICABILITY" Then coverage = coverage + 1
 Next
 Check coverage = 1, "Group omitted member coverage gap"
 Set preview = NAVFITGroupPreview(groupID): Set result = preview("Readiness"): coverage = 0
 For Each item In result
  If item("ReportID") = id And item("Code") = "VR-APPLICABILITY" Then coverage = coverage + 1
 Next
 Check coverage = 1, "Group export preflight omitted member coverage gap"
 Set pending = SavedFields(id): pending("PeriodStart") = #1/1/2026#: pending("PeriodEnd") = #9/30/2026#
 Set overrides = CreateObject("Scripting.Dictionary"): overrides.Add CStr(id), pending
 revisions = DCount("*", "ReportRevision"): Set result = ValidateLiveGroup(groupID, overrides): coverage = 0
 For Each item In result
  If item("ReportID") = id And item("Code") = "VR-APPLICABILITY" Then coverage = coverage + 1
 Next
 Check coverage = 0, "Group ignored pending applicability override"
 Set pending = SavedFields(id)
 Check DCount("*", "ReportRevision") = revisions And Year(pending("PeriodEnd")) = 2024, "Pending validation wrote report values/history"
 ' Readiness runs before the extension guard; no file or export operation is created.
 On Error Resume Next
 ignored = NAVFITExportReport(id, CurrentProject.Path & "\c04-preflight.invalid", "")
 exportError = Err.Number: Err.Clear
 On Error GoTo 0
 Check exportError = vbObjectError + 5300, "Single export did not reach extension guard"
 Set result = NAVFITLastReportReadiness()
 Check Has(result, "VR-APPLICABILITY"), "Single-report export omitted historical coverage"
 changes.RemoveAll: changes("PeriodStart") = original("PeriodStart"): changes("PeriodEnd") = original("PeriodEnd")
 CommitSavedFields id, CurrentRevision(id), changes, "Restore current test period"
End Sub
Public Function TestReadiness() As String
 Dim stage As String, id As Long, kind As Variant, f As Object, result As Collection, key As Variant, changes As Object, pack As Long, legacy As Long, run As Long
 Dim groupID As Long, other As Long, preview As Object, rows As Long, before As Long, item As Object, db As DAO.Database, member As Long
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC READINESS", "Tester"
 id = CommitDraft(0, "EVAL", "Synthetic", "Example", Null, Null, "")
 For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
  stage = CStr(kind): Set f = Example(id, CStr(kind)): Set result = ValidateLiveFields(f)
  Check result.Count = 0, kind & " valid baseline has " & result.Count & " findings: " & ReadinessSummary(result)
  f("CommandEmployment") = Null: f("DutyDescription") = Null: f("ReadinessSummary") = Null
  Set result = ValidateLiveFields(f): Check Has(result, "VR-DUTIES", "CommandEmployment"), "Block28 missing not found"
  Check Has(result, "VR-DUTIES", "DutyDescription") And Has(result, "VR-PFA"), "Duties/readiness missing not found"
  For Each item In result: Check Len(item("Reference")) > 20, "Finding lacks reference": Next
  Set f = Example(id, CStr(kind)): f("ReportRegular") = "No": f("ReportConcurrent") = "Yes": f("CounselingDate") = Null: f("CounselorName") = Null: f("ReadinessSummary") = Null: f("ReadinessDate") = Null
  Set result = ValidateLiveFields(f): Check Not Has(result, "COUNSELING") And Not Has(result, "VR-PFA"), "Pure Concurrent blank exception rejected"
  f("ReadinessDate") = "2026-06-01": Set result = ValidateLiveFields(f): Check Has(result, "VR-PFA"), "Concurrent partial PFA entry accepted"
  f("ReadinessDate") = Null
  f("ReportRegular") = "Yes": Set result = ValidateLiveFields(f): Check Has(result, "COUNSELING") And Has(result, "VR-PFA"), "Concurrent/Regular skipped required entries"
  Set f = Example(id, CStr(kind)): f("ReportRegular") = "No": f("ReportOperationalCommander") = "Yes"
  Set result = ValidateLiveFields(f): Check Has(result, "VR-OPS-COUNSEL"), "OPS counseling not detected"
  f("CounselingDate") = Null: f("CounselorName") = Null: f("ReadinessSummary") = "N"
  Set result = ValidateLiveFields(f): Check Not Has(result, "COUNSELING") And Not Has(result, "VR-OPS-COUNSEL"), "OPS blank counseling rejected"
  Set f = Example(id, CStr(kind)): f("Trait33") = "1.0": f("OfficialComments") = Null
  Set result = ValidateLiveFields(f): Check Has(result, "JUSTIFY-TRAITS", "OfficialComments"), "Required-empty adverse comments missed"
  f("MemberSSN") = "000000000": Set result = ValidateLiveFields(f): Check Has(result, "NV-SSN"), "Placeholder SSN accepted"
 Next
 stage = "Relationships": Set f = Example(id, "EVAL")
 f("AssignmentRecommendationA") = Null: f("AssignmentRecommendationB") = "TEST": f("SeniorRaterName") = f("RaterName"): f("RetentionYes") = "No"
 Set result = ValidateLiveFields(f): Check Has(result, "CAREER-PRESENT") And Has(result, "NV-RATER") And Has(result, "VR-RETENTION"), "Relationship findings missing"
 f("ReportingSeniorDesignator") = "ABC": f("ReportingSeniorAddress1") = Null: f("ReportingSeniorGrade") = "CDR"
 Set result = ValidateLiveFields(f): Check Has(result, "VR-RS-DESIG") And Has(result, "VR-CONTACT"), "Senior compatibility/contact missing"
 f("RankRate") = "ITC": Set result = ValidateLiveFields(f): Check Has(result, "NV-RATE"), "Report type/grade mismatch missed"
 stage = "Effective dated schedules"
 CheckApplicability id
 CheckHistoricalPack id
 Check PeriodicExpectedDate("O6", #7/31/2026#) = #7/31/2026#, "2026 O6 schedule"
 Check PeriodicExpectedDate("O6", #3/31/2027#) = #3/31/2027#, "2027 O6 schedule"
 Check PeriodicExpectedDate("O5", #4/30/2026#) = #4/30/2026# And PeriodicExpectedDate("O5", #6/30/2027#) = #6/30/2027#, "O5 schedule transition"
 Check PeriodicExpectedDate("O2", #2/29/2028#) = #2/29/2028#, "Leap-year schedule"
 Check PeriodicExpectedDate("E8", #9/15/2026#, "CMDCS") = #11/15/2026#, "CMDCS exception"
 Set f = Example(id, "EVAL"): f("OccasionPeriodic") = "Yes": f("PeriodEnd") = #9/30/2026#
 Set result = ValidateLiveFields(f): Check Has(result, "VR-PERIODIC"), "Periodic mismatch not found"
 f("RankRate") = "MCPON": Set result = ValidateLiveFields(f): Check Has(result, "VR-PERIODIC"), "Exempt/unknown grade not reviewed"
 stage = "Saved parity"
 Set f = Example(id, "EVAL"): Set changes = CreateObject("Scripting.Dictionary")
 For Each key In DraftFieldKeys(): If key <> "FormKind" Then changes(CStr(key)) = f(CStr(key))
 Next
 CommitSavedFields id, CurrentRevision(id), changes, "Synthetic valid fields"
 pack = InstallNavyDraftChecks(): Check InstallNavyDraftChecks() = pack, "Shared pack not idempotent"
 Check IsNull(MajorScalar("SELECT EffectiveTo FROM RulePack WHERE PackID=" & pack)), "Shared pack has stale cutoff"
 run = RunRulePack(id, pack): Check DCount("*", "ValidationFinding", "ValidationRunID=" & run & " AND FindingState='Pass'") = 1, "Saved valid result differs from live"
 changes.RemoveAll: changes("CommandEmployment") = Null: changes("ReadinessSummary") = Null
 CommitSavedFields id, CurrentRevision(id), changes, "Synthetic incomplete draft"
 Set result = ValidateLiveFields(SavedFields(id), id): run = RunRulePack(id, pack)
 Check DCount("*", "ValidationFinding", "ValidationRunID=" & run) = result.Count, "Saved/live finding counts diverge"
 legacy = InstallLegacyNavyDraftChecks(): Check legacy <> pack, "Legacy history overwritten"
 before = DCount("*", "ValidationRun"): Set db = CurrentDb
 db.TableDefs("ValidationFinding").ValidationRule = "[FindingState]<>'Fail'"
 Check SavedRejected(id, pack), "Injected saved-finding failure accepted"
 Check DCount("*", "ValidationRun") = before, "Failed saved run left partial state"
 db.TableDefs("ValidationFinding").ValidationRule = ""
 Check RunRulePack(id, pack) > run, "Saved readiness retry failed"
 stage = "Export preflight"
 groupID = CreateSummaryGroup("Synthetic E6", "EVAL", "E6"): AddGroupMember groupID, id, "E6"
 Set preview = NAVFITGroupPreview(groupID): Set result = preview("Readiness")
 Check Has(result, "VR-DUTIES") And InStr(preview("Text"), "Readiness:") > 0, "Group export omitted shared findings"
 Check UXOpenNAVFITGroupExport(groupID), "Export view failed"
 Check UXExportReadiness(True), "Full group issues not available"
 Check Forms!frmUXValidation.OpenArgs = "G:" & groupID, "Wrong findings context"
 DoCmd.Close acForm, "frmUXValidation", acSaveNo
 stage = "Historical applicability paths": CheckReadinessPaths id, groupID
 other = DuplicateDraft(id): AddGroupMember groupID, other, "E6"
 changes.RemoveAll: changes("ReportingSeniorGrade") = "LCDR": changes("OccasionPeriodic") = "Yes"
 CommitSavedFields other, CurrentRevision(other), changes, "Synthetic group differences"
 Set result = ValidateLiveGroup(groupID)
 Check Has(result, "NV-RS-GROUP", "ReportingSeniorGrade"), "Live group omitted RS grade difference"
 Check Has(result, "NV-GROUP", "OccasionPeriodic"), "Native group occasion difference missed"
 stage = "Known Regular continuity"
 member = MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & id)
 db.Execute "UPDATE MemberReport SET MemberID=" & member & ",MatchState='Confirmed' WHERE ReportID In (" & id & "," & other & ")", dbFailOnError
 changes.RemoveAll: changes("PeriodStart") = #1/1/2025#: changes("PeriodEnd") = #12/30/2025#
 CommitSavedFields other, CurrentRevision(other), changes, "Synthetic prior Regular"
 Set result = ValidateLiveFields(SavedFields(id), id): Check Has(result, "VR-CONTINUITY"), "Known Regular gap missed"
 TestReadiness = "PASS: " & n & " shared readiness assertions": Exit Function
Failed:
 TestReadiness = "FAIL at " & stage & ": " & Err.Number & " " & Err.Description
End Function
