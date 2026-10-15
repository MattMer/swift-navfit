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
