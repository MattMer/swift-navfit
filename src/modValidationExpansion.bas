Attribute VB_Name = "modValidationExpansion"
Option Compare Binary
Option Explicit
' One shared finding catalogue supplements live, saved and export readiness.
Public Function VRText(ByVal f As Object, ByVal key As String) As String
 If f.Exists(key) Then VRText = Trim$(Nz(f(key), ""))
End Function
Public Function RequiresCounselingEntry(ByVal f As Object) As Boolean
 If VRText(f, "ReportOperationalCommander") = "Yes" Then Exit Function
 If VRText(f, "ReportConcurrent") = "Yes" And VRText(f, "ReportRegular") <> "Yes" Then
  RequiresCounselingEntry = (VRText(f, "CounselingDate") <> "" Or VRText(f, "CounselorName") <> "")
 Else
  RequiresCounselingEntry = True
 End If
End Function
Private Function VRIn(ByVal value As String, ByVal values As String) As Boolean
 VRIn = (InStr(1, "|" & values & "|", "|" & UCase$(value) & "|", vbBinaryCompare) > 0)
End Function
' These are compatibility warnings, not a claim that native legacy rules are policy.
Private Sub VRWarn(ByVal findings As Collection, ByVal f As Object, ByVal id As Long, ByVal key As String, ByVal code As String, ByVal message As String)
 VAdd findings, f, id, key, "Needs review", code, message
End Sub
Public Sub ExpandReportValidation(ByVal findings As Collection, ByVal f As Object, ByVal id As Long)
 Dim key As Variant, kind As String, text As String, grade As String, desig As String, mark As String, adverse As Boolean
 kind = VRText(f, "FormKind")
 For Each key In Array("MemberSSN", "ReportingSeniorSSN")
  If VRText(f, CStr(key)) = "000000000" Then VAdd findings, f, id, CStr(key), "Incomplete draft", "NV-SSN", "All-zero SSN is a placeholder. Replace it before native validation; synthetic drafts may still be saved."
 Next
 For Each key In Array("CommandEmployment", "PrimaryDuty", "DutyDescription")
  If VRText(f, CStr(key)) = "" Then VAdd findings, f, id, CStr(key), "Incomplete draft", "VR-DUTIES", "Complete this command employment/duties block before final review."
 Next
 If kind <> "CHIEFEVAL" Then
  If VRText(f, "AssignmentRecommendationB") <> "" And VRText(f, "AssignmentRecommendationA") = "" Then VAdd findings, f, id, "AssignmentRecommendationA", "Incomplete draft", "CAREER-PRESENT", "A second career recommendation needs a first recommendation."
 End If
 If kind = "EVAL" Then
  If VRText(f, "RetentionYes") <> "Yes" And VRText(f, "RetentionNo") <> "Yes" Then VAdd findings, f, id, "RetentionYes", "Incomplete draft", "VR-RETENTION", "Select Recommended or Not Recommended for retention."
  If VRText(f, "ObservationStatus") <> "NOB" Then
   For Each key In Array("RaterName", "SeniorRaterName")
    text = VRText(f, CStr(key))
    If text = "" Then
     VAdd findings, f, id, CStr(key), "Incomplete draft", "VR-RATER", "Enter the applicable rater; NONE AVAILABLE is permitted for senior rater when appropriate."
    ElseIf Not (key = "SeniorRaterName" And UCase$(text) = "NONE AVAILABLE") Then
     If InStr(text, ",") = 0 Then VRWarn findings, f, id, CStr(key), "NV-RATER", "Check name, initials and grade/rate formatting for NAVFIT98A."
    End If
   Next
  End If
  text = VRText(f, "RaterName")
  If text <> "" And UCase$(text) <> "NONE AVAILABLE" And StrComp(text, VRText(f, "SeniorRaterName"), vbTextCompare) = 0 Then VRWarn findings, f, id, "SeniorRaterName", "NV-RATER", "Rater and senior rater are identical. Verify the roles and any permitted exception."
 End If
 For Each key In TraitKeys()
  mark = VRText(f, CStr(key))
  If mark = "1.0" Then adverse = True
 Next
 If VRText(f, "PromotionRecommendation") = "Significant Problems" Or VRText(f, "RetentionNo") = "Yes" Then adverse = True
 For Each key In findings
  If key("Code") = "JUSTIFY-TRAITS" Then adverse = True
 Next
 If adverse And VRText(f, "OfficialComments") = "" Then VAdd findings, f, id, "OfficialComments", "Incomplete draft", "JUSTIFY-TRAITS", "Required adverse-report justification is empty. Text presence alone will not establish adequate justification."
 If VRText(f, "ReportOperationalCommander") = "Yes" Then
  If VRText(f, "CounselingDate") <> "" Or VRText(f, "CounselorName") <> "" Then VAdd findings, f, id, "CounselingDate", "Policy issue", "VR-OPS-COUNSEL", "Operational Commander counseling blocks must be blank."
 ElseIf UCase$(VRText(f, "CounselingDate")) = "NOT REQ" And VRText(f, "CounselorName") <> "" Then
  VRWarn findings, f, id, "CounselorName", "NV-COUNSEL", "NOT REQ has a counselor entry. Verify the counseling block and native compatibility; do not erase pending text automatically."
 End If
 VRReadiness findings, f, id
 VRPeople findings, f, id
 VRPeriodic findings, f, id
 VRContinuity findings, f, id
 If VRText(f, "ReportConcurrent") = "Yes" Then VRWarn findings, f, id, "ReportConcurrent", "NV-RRS", "Verify the separate regular reporting senior name, grade, command and UIC on the native signature block. Native-only imported details are retained; these are not separate editable Swift fields."
End Sub
' Apply form-specific PFA coding and the pure Concurrent exception.
Private Sub VRReadiness(ByVal findings As Collection, ByVal f As Object, ByVal id As Long)
 Dim text As String, dateText As String, i As Long
 text = UCase$(VRText(f, "ReadinessSummary")): dateText = VRText(f, "ReadinessDate")
 If text = "" And dateText = "" And VRText(f, "ReportConcurrent") = "Yes" And VRText(f, "ReportRegular") <> "Yes" Then Exit Sub
 If text = "" Then
  VAdd findings, f, id, "ReadinessSummary", "Incomplete draft", "VR-PFA", "Complete physical readiness. Use the form's no-test convention when no official results apply."
 ElseIf VRText(f, "FormKind") <> "CHIEFEVAL" Then
  For i = 1 To Len(text)
   If InStr("PBFMWN", Mid$(text, i, 1)) = 0 Then VAdd findings, f, id, "ReadinessSummary", "Policy issue", "VR-PFA", "Use P/B/F/M/W/N codes, oldest to newest, without spaces or slashes.": Exit For
  Next
  If InStr(text, "B") > 0 And VRText(f, "DutyDescription") = "" Then VAdd findings, f, id, "DutyDescription", "Incomplete draft", "VR-PFA", "A B physical-readiness result needs the block 29 explanation."
 Else
  If dateText = "" And text <> "N" And text <> "OTHER" Then VRWarn findings, f, id, "ReadinessDate", "NV-PFA", "CHIEFEVAL: verify the most recent official test date/result, or Other with a block 29 explanation when no test applies."
  If (text = "N" Or text = "OTHER") And VRText(f, "DutyDescription") = "" Then VAdd findings, f, id, "DutyDescription", "Incomplete draft", "VR-PFA", "Explain the CHIEFEVAL no-test/Other result in block 29."
 End If
End Sub
' Versioned calendars from CH-1 enclosure (1), tables 1/1A in the retained
' CH-2 compilation (PDF17/18). Earlier calendars have not been verified.
Public Function PeriodicScheduleVersion(ByVal ending As Date) As String
 Dim row As Variant
 For Each row In Array(Array(#12/16/2025#, #12/31/2026#, "CH1-T1"), Array(#1/1/2027#, #12/31/9999#, "CH1-T1A"))
  If DateValue(ending) >= row(0) And DateValue(ending) <= row(1) Then PeriodicScheduleVersion = row(2): Exit Function
 Next
End Function
Public Function ValidationApplicabilityReason(ByVal f As Object) As String
 Dim row As Variant, ending As Date
 If Not IsDate(VRText(f, "PeriodEnd")) Then
  ValidationApplicabilityReason = "An ending date is required to select policy applicability."
  Exit Function
 End If
 ending = DateValue(f("PeriodEnd"))
 ' General-rule coverage is distinct from the calendar and exemption windows.
 For Each row In Array(Array(#5/26/2026#, #12/31/9999#, "CH2-GENERAL"))
  If ending >= row(0) And ending <= row(1) Then Exit Function
 Next
 ValidationApplicabilityReason = "General policy checks use the verified CH-2 compilation from 26 May 2026; earlier policy coverage is not verified. Periodic calendars are evaluated separately from 16 December 2025."
End Function
Public Function PeriodicPolicyReference(ByVal f As Object) As String
 Dim ending As String
 ending = VRText(f, "PeriodEnd")
 PeriodicPolicyReference = "BUPERSINST 1610.10H CH-1, 16 December 2025, enclosure (1), table 1 (through 2026), table 1A (from 2027); retained CH-2 compilation PDF pages 17-18. Source SHA-256: " & NAVY_RULE_SOURCE_HASH
 If IsDate(ending) Then PeriodicPolicyReference = PeriodicPolicyReference & vbCrLf & "Selected calendar: " & PeriodicScheduleVersion(CDate(ending))
 PeriodicPolicyReference = PeriodicPolicyReference & vbCrLf & "Command-senior-enlisted exceptions: CH-2 enclosure (2), paragraph 3-4a-d, 26 May 2026. Highest-paygrade Regular periodic exemptions: NAVADMIN 132/26, paragraphs 1 and 3, effective 3 June 2026 (authoritative effective date). Permanent LDO grade/NEC/assignment require human confirmation."
 PeriodicPolicyReference = PeriodicPolicyReference & vbCrLf & "https://www.mynavyhr.navy.mil/Portals/55/Messages/NAVADMIN/NAV2026/NAV26132.pdf"
End Function
Public Function PeriodicExpectedDate(ByVal worn As String, ByVal ending As Date, Optional ByVal rankRate As String = "") As Variant
 Dim version As String, entries As String, entry As Variant, parts As Variant, month As Long, dayNumber As Long
 PeriodicExpectedDate = Null
 version = PeriodicScheduleVersion(ending): If version = "" Then Exit Function
 ' Month/day=0 means the last day of that month; handles leap/century years.
 entries = "O1:5:0|O2:2:0|O3:1:0|O4:10:0|W1:9:0|W2:9:0|W3:3:0|W4:3:0|W5:3:0|E1:7:15|E2:7:15|E3:7:15|E4:6:15|E5:3:15|E6:11:15|E7:9:15|E8:9:15|E9:4:15"
 If version = "CH1-T1" Then entries = entries & "|O5:4:0|O6:7:0" Else entries = entries & "|O5:6:0|O6:3:0"
 If UCase$(Trim$(rankRate)) = "CMDCS" Then
  If UCase$(worn) <> "E8" Then Exit Function
  entries = "E8:11:15"
 End If
 For Each entry In Split(entries, "|")
  parts = Split(CStr(entry), ":")
  If UCase$(worn) = parts(0) Then
   month = CLng(parts(1)): dayNumber = CLng(parts(2))
   If dayNumber = 0 Then
    dayNumber = Day(DateSerial(Year(ending), month + 1, 0))
   End If
   PeriodicExpectedDate = DateSerial(Year(ending), month, dayNumber): Exit Function
  End If
 Next
End Function
' Warn only from explicitly linked, known Regular reports; prior imported evals
' do not record report type, so they cannot prove a Regular continuity gap.
Private Sub VRContinuity(ByVal findings As Collection, ByVal f As Object, ByVal id As Long)
 Dim member As Long, previous As Variant, startText As String, db As DAO.Database, r As DAO.Recordset
 If id = 0 Or VRText(f, "ReportRegular") <> "Yes" Then Exit Sub
 startText = VRText(f, "PeriodStart"): If Not IsDate(startText) Then Exit Sub
 member = Nz(MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & id & " AND MatchState='Confirmed'"), 0)
 If member = 0 Then Exit Sub
 Set db = CurrentDb
 Set r = db.OpenRecordset("SELECT Max(r.PeriodEnd) AS PriorEnd FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & member & " AND m.MatchState='Confirmed' AND r.IsArchived=False AND r.ReportID<>" & id & " AND r.ReportRegular='Yes' AND r.FormKind=" & QuoteSQL(VRText(f, "FormKind")) & " AND r.PeriodStart<#" & Format$(CDate(startText), "yyyy-mm-dd") & "#", dbOpenSnapshot)
 previous = r!PriorEnd.Value: r.Close
 If Not IsNull(previous) Then
  If CDate(startText) <> DateAdd("d", 1, CDate(previous)) Then VRWarn findings, f, id, "PeriodStart", "VR-CONTINUITY", "Known prior Regular report ends " & Format$(previous, "yyyy-mm-dd") & "; verify a possible gap/overlap. Local history may be incomplete or include alternative drafts."
 End If
End Sub
Private Sub VRPeriodic(ByVal findings As Collection, ByVal f As Object, ByVal id As Long)
 Dim ending As String, worn As String, expected As Variant
 If VRText(f, "OccasionPeriodic") <> "Yes" Then Exit Sub
 ending = VRText(f, "PeriodEnd"): If Not IsDate(ending) Then Exit Sub
 worn = ValidationWornPaygrade(VRText(f, "RankRate"))
 If CDate(ending) >= #6/3/2026# And VRText(f, "ReportRegular") = "Yes" And (worn = "E9" Or worn = "W5" Or (worn = "O6" And Left$(VRText(f, "Designator"), 1) = "6")) Then
  VRWarn findings, f, id, "OccasionPeriodic", "VR-PERIODIC", "NAVADMIN 132/26 removes Regular periodic reporting from 3 June 2026 for E9, CWO5 and permanent LDO O6. Confirm eligibility/permanent grade; detachment requirements remain separate. No report or history is deleted."
  Exit Sub
 End If
 If CDate(ending) >= #5/26/2026# And UCase$(VRText(f, "RankRate")) = "CMDCM" Then
  VRWarn findings, f, id, "OccasionPeriodic", "VR-PERIODIC", "Verify the CH-2 paragraph 3-4 command-master-chief exception from 26 May 2026; 8CMC NEC and flag/general-officer reporting relationship are not established by the rate label alone."
  Exit Sub
 End If
 expected = PeriodicExpectedDate(worn, CDate(ending), VRText(f, "RankRate"))
 If IsNull(expected) Then
  VRWarn findings, f, id, "PeriodEnd", "VR-PERIODIC", "No verified periodic schedule for this grade/date; review the applicable instruction."
 ElseIf DateValue(ending) <> CDate(expected) Then
  VRWarn findings, f, id, "PeriodEnd", "VR-PERIODIC", "Scheduled periodic ending date: " & Format$(expected, "yyyy-mm-dd") & ". Verify exception, recent observed report, command-senior-enlisted assignment and applicable table before changing dates."
 End If
End Sub
' Reuse rows already collected by group validation; no per-rule requery.
Public Sub ExpandGroupValidation(ByVal findings As Collection, ByVal rows As Collection)
 Dim row As Object, f As Object, key As Variant, first As String, different As Boolean, item As Object
 For Each key In Array("OccasionPeriodic", "OccasionSpecial")
  first = "@": different = False
  For Each row In rows
   Set f = row("Fields")
   If first = "@" Then first = VRText(f, CStr(key)) Else If first <> VRText(f, CStr(key)) Then different = True
  Next
  If different Then
   For Each row In rows
    Set f = row("Fields"): VRWarn findings, f, row("ID"), CStr(key), "NV-GROUP", "Report occasion flags differ within this group. NAVFIT98A checks consistency; verify group membership and report occasions."
   Next
  End If
 Next
End Sub
' Format a bounded preview; the issues form provides every finding and full citations.
Public Function ReadinessSummary(ByVal findings As Collection) As String
 Dim item As Object, text As String, i As Long
 text = "Readiness: " & findings.Count & " findings. Incomplete drafts remain saveable; native validation is still required."
 For Each item In findings
  i = i + 1
  If i <= 25 Then text = text & vbCrLf & item("Member") & " | " & item("Label") & " | " & item("Severity") & ": " & item("Message")
 Next
 If findings.Count > 25 Then text = text & vbCrLf & "Showing 25 findings. Open All readiness issues for the complete list and policy references."
 ReadinessSummary = text
End Function
' Read back staged native totals before publication; mismatches fail the atomic export.
Public Sub VerifyNativeGroupSummary(ByVal db As DAO.Database, ByVal preview As Object)
 Dim r As DAO.Recordset, key As Variant, i As Long
 Set r = db.OpenRecordset("SELECT SummarySP,SummaryProg,SummaryProm,SummaryMP,SummaryEP FROM Reports WHERE ReportType Is Not Null UNION ALL SELECT SummarySP,SummaryProg,SummaryProm,SummaryMP,SummaryEP FROM Folders WHERE FolderID=1", dbOpenSnapshot)
 Do While Not r.EOF
  i = 1
  For Each key In Array("SummarySP", "SummaryProg", "SummaryProm", "SummaryMP", "SummaryEP")
   If Nz(r.Fields(CStr(key)), "") <> CStr(preview("Count" & i)) Then r.Close: Err.Raise vbObjectError + 8410, , "Native summary counts differ from the exported saved roster. Nothing published."
   i = i + 1
  Next
  r.MoveNext
 Loop
 r.Close
End Sub
' Validate visible representations. Never require hidden name parts or silently rewrite names.
Private Sub VRPeople(ByVal findings As Collection, ByVal f As Object, ByVal id As Long)
 Dim text As String, grade As String, desig As String, key As Variant, worn As String
 text = VRText(f, "ReportingSeniorName")
 If text <> "" And InStr(text, ",") = 0 Then VRWarn findings, f, id, "ReportingSeniorName", "VR-RS-NAME", "Check reporting senior surname, comma and initials; verify the actual name and signature."
 grade = UCase$(VRText(f, "ReportingSeniorGrade")): desig = UCase$(VRText(f, "ReportingSeniorDesignator"))
 If grade <> "" Then
  If VRIn(grade, "ADM|VADM|RADM|RDML|CAPT|CDR|LCDR|LT|LTJG|ENS") Then
   If Not (Len(desig) = 4 And desig Like "####") And desig <> "LTR" And Not VRIn(desig, "USCG|USPH|NOAA") Then VRWarn findings, f, id, "ReportingSeniorDesignator", "VR-RS-DESIG", "Verify naval officer designator, permitted other-service branch or LTR authorization."
  ElseIf VRIn(grade, "GEN|LGEN|MGEN|BGEN|COL|LTC|MAJ|CPT|1LT|2LT") Then
   If Not VRIn(desig, "USA|USAF|USMC|USSF|LTR") Then VRWarn findings, f, id, "ReportingSeniorDesignator", "VR-RS-DESIG", "Verify other-service reporting senior grade and branch designation."
  ElseIf grade = "CIV" Or grade Like "GS*" Or grade Like "SES*" Or Left$(ValidationWornPaygrade(grade), 1) = "E" Then
   If desig <> "" And desig <> "LTR" Then VRWarn findings, f, id, "ReportingSeniorDesignator", "VR-RS-DESIG", "Enlisted/civilian reporting senior designator is normally blank; verify LTR authorization if used."
  ElseIf Left$(grade, 3) = "CWO" Or grade = "WO1" Then
   If Not (Len(desig) = 4 And desig Like "####") And Not VRIn(desig, "USA|USAF|USMC|USSF|USCG|LTR") Then VRWarn findings, f, id, "ReportingSeniorDesignator", "VR-RS-DESIG", "Verify warrant officer designator or service branch."
  Else
   VRWarn findings, f, id, "ReportingSeniorGrade", "VR-RS-DESIG", "Reporting senior grade is outside the supported naval, other-service and civilian formats; verify rather than guess."
  End If
 End If
 For Each key In Array("ReportingSeniorAddress1", "ReportingSeniorPhone")
  If VRText(f, CStr(key)) = "" Then VAdd findings, f, id, CStr(key), "Incomplete draft", "VR-CONTACT", "Complete the reporting senior address/phone for the OMPF copy; member-copy address omissions have exceptions."
 Next
 If VRText(f, "FormKind") = "CHIEFEVAL" Then
  For Each key In Array("ReportingSeniorCity", "ReportingSeniorState", "ReportingSeniorZip")
   If VRText(f, CStr(key)) = "" Then VRWarn findings, f, id, CStr(key), "NV-CONTACT", "Verify structured CHIEFEVAL address completeness, including overseas/APO/FPO exceptions."
  Next
 End If
 worn = ValidationWornPaygrade(VRText(f, "RankRate"))
 If worn <> "" Then
  If (VRText(f, "FormKind") = "EVAL" And Not VRIn(worn, "E1|E2|E3|E4|E5|E6")) Or (VRText(f, "FormKind") = "CHIEFEVAL" And Not VRIn(worn, "E7|E8|E9")) Or (VRText(f, "FormKind") = "FITREP" And Left$(worn, 1) = "E") Then VRWarn findings, f, id, "RankRate", "NV-RATE", "Worn grade is incompatible with the selected NAVPERS report type."
 ElseIf VRText(f, "RankRate") <> "" Then
  VRWarn findings, f, id, "RankRate", "NV-RATE", "Worn grade/rate was not recognized; verify the current rating code."
 End If
 text = VRText(f, "Designator")
 If VRText(f, "FormKind") = "FITREP" Then
  If Len(text) <> 4 Or Not text Like "####" Then VRWarn findings, f, id, "Designator", "NV-DESIG", "Verify the four-digit officer designator. Unknown or retired codes need review."
 ElseIf text <> "" Then
  If InStr(text, " ") > 0 Or UBound(Split(text, "/")) > 1 Then VRWarn findings, f, id, "Designator", "NV-DESIG", "Verify at most two warfare designations, separated by a slash and without spaces."
 End If
 text = UCase$(VRText(f, "BilletSubcategory"))
 If text <> "" And Not VRIn(text, "NA|BASIC|APPROVED|INDIV AUG|CO AFLOAT|CO ASHORE|OIC|SEA COMP|CRF|CANVASSER|RESIDENT|INTERN|INSTRUCTOR|STUDENT|RESAC1|RESAC6|SCREENED") Then VRWarn findings, f, id, "BilletSubcategory", "NV-SUBCAT", "Verify this unrecognized or specially authorized billet code and command eligibility against table 1-1."
 If text = "SCREENED" And Not VRIn(worn, "E7|E8") Then VRWarn findings, f, id, "BilletSubcategory", "NV-SUBCAT", "SCREENED normally applies to E7/E8; verify the current screening and worn grade."
End Sub
