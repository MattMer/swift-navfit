Attribute VB_Name = "modLiveValidation"
Option Compare Binary
Option Explicit
Private Function VText(ByVal f As Object, ByVal key As String) As String
    If f.Exists(key) Then VText = Trim$(Nz(f(key), ""))
End Function
Public Function ValidationPolicyReference(ByVal code As String) As String
    Dim page As String, locator As String, excerpt As String, pdfPage As Long
    Select Case code
        Case "NAME-PRESENT": page = "1-1": pdfPage = 30: locator = "paragraph 1-2, NAME, block 1": excerpt = "Enter the member's last name, first name, middle initial, and suffix (if any)"
        Case "RATE-PRESENT": page = "1-1": pdfPage = 30: locator = "paragraph 1-2, GRADE/RATE, block 2": excerpt = "Enter the grade or rate the Service member is actually wearing"
        Case "SSN-PRESENT": page = "1-2": pdfPage = 31: locator = "paragraph 1-2, SSN, block 4": excerpt = "Enter the Service member's full SSN with hyphens after the third and fifth digits."
        Case "UIC-PRESENT": page = "1-2": pdfPage = 31: locator = "paragraph 1-2, Unit ID Code, block 6": excerpt = "Enter the UIC of the Service member's ship or station."
        Case "COMPONENT": page = "1-2": pdfPage = 31: locator = "paragraph 1-2, DUTY/COMPETITIVE STATUS, block 5": excerpt = "Place an X in only one box"
        Case "REPORT-TYPE": page = "1-5": pdfPage = 34: locator = "paragraph 1-2, TYPE OF REPORT, blocks 17-19": excerpt = "If this is a Concurrent/Regular report, place an ""X"" in blocks 17 and 18."
        Case "WHOLE-NOB": page = "1-5": pdfPage = 34: locator = "paragraph 1-2, NOT OBSERVED REPORT, block 16; exceptions in chapters 6 and 10": excerpt = "The trait grades must be left blank, no career recommendation is permitted"
        Case "TRAIT-RECOMMENDATION": page = "1-16": pdfPage = 45: locator = "paragraph 1-2, INDIVIDUAL (PROMOTION RECOMMENDATION)": excerpt = "A Promotable or higher recommendation may not be assigned with any trait graded 1.0."
        Case "TRAIT-TWO": page = "1-16": pdfPage = 45: locator = "paragraph 1-2, INDIVIDUAL (PROMOTION RECOMMENDATION)": excerpt = "A recommendation of Must Promote or Early Promote may not be assigned with any trait assessed as 2.0."
        Case "JUNIOR-OFFICER": page = "1-16": pdfPage = 45: locator = "paragraph 1-2, promotion recommendation, Note 2": excerpt = "are prohibited from receiving a promotion recommendation higher than Promotable."
        Case "OCCASION": page = "1-4": pdfPage = 33: locator = "paragraph 1-2, OCCASION FOR REPORT, blocks 10-13, Note": excerpt = "More than one occasion may apply, except Special cannot be combined with another occasion."
        Case "RS-SSN": page = "1-11": pdfPage = 40: locator = "paragraph 1-2, reporting senior SSN, block 27": excerpt = "Enter the reporting senior's full SSN with hyphens after the third and fifth digits"
        Case "PERIOD": page = "1-4": pdfPage = 33: locator = "paragraph 1-2, PERIOD OF REPORT, blocks 14-15": excerpt = "Enter the beginning FROM and ending TO dates of the report period."
        Case "TRAIT-CRITICAL": page = "1-16": pdfPage = 45: locator = "paragraph 1-2, promotion recommendation, observed trait restrictions": excerpt = "Accountability (CHIEFEVAL) must be evaluated as 3.0 or higher"
        Case "TRAIT-TWO-COUNT": page = "1-16": pdfPage = 45: locator = "paragraph 1-2, promotion recommendation, observed trait restrictions": excerpt = "A Promotable promotion recommendation allows up to two traits, excluding Character or Equal Opportunity"
        Case "PROMOTION-PRESENT", "PARTIAL-NOB": page = "1-16": pdfPage = 45: locator = "paragraph 1-2, INDIVIDUAL (PROMOTION RECOMMENDATION)": excerpt = "Ensure justification is provided in the comments section if a partially-graded NOB report is issued."
        Case "GROUP-OFFICER": page = "1-19": pdfPage = 48: locator = "paragraph 1-2, SUMMARY, Table 1-3": excerpt = "Group officer reports that share all the following characteristics."
        Case "GROUP-ENLISTED": page = "1-22": pdfPage = 51: locator = "paragraph 1-2, SUMMARY, Table 1-4": excerpt = "Group enlisted reports that share all the following characteristics."
        Case "GROUP-CATEGORY": page = "1-20 to 1-21": pdfPage = 49: locator = "paragraph 1-2, Officer Competitive Categories": excerpt = "Each category consists of all designators within the parentheses."
        Case "CAPACITY": page = "1-17 to 1-18": pdfPage = 46: locator = "paragraph 1-2, Upper Limits; Table 1-2": excerpt = "Twenty percent of each summary group (rounded up to nearest whole number)."
        Case "SHIP-PRESENT": page = "1-3": pdfPage = 32: locator = "paragraph 1-2, SHIP/STATION, block 7": excerpt = "Enter the abbreviated name of the activity to which the member is assigned"
        Case "STATUS-PRESENT": page = "1-3": pdfPage = 32: locator = "paragraph 1-2, PROMOTION STATUS, block 8": excerpt = "Enter the member's promotion status on the ending date of the report period"
        Case "REPORTED-PRESENT": page = "1-4": pdfPage = 33: locator = "paragraph 1-2, DATE REPORTED, block 9": excerpt = "Use the date the Service member reported to the command in block 7."
        Case "RS-NAME": page = "1-9": pdfPage = 38: locator = "paragraph 1-2, REPORTING SENIOR, block 22": excerpt = "Enter the last name, initials, and suffix (if any) of the reporting senior."
        Case "RS-GRADE": page = "1-10": pdfPage = 39: locator = "paragraph 1-2, GRADE, block 23": excerpt = "Enter the reporting senior's grade."
        Case "RS-TITLE": page = "1-10": pdfPage = 39: locator = "paragraph 1-2, TITLE, block 25": excerpt = "title (abbreviated if necessary)."
        Case "RS-UIC": page = "1-11": pdfPage = 40: locator = "paragraph 1-2, UIC, block 26": excerpt = "Enter the UIC of the reporting senior's command."
        Case "UIC-FORMAT": page = "1-2": pdfPage = 31: locator = "paragraph 1-2, Unit ID Code, block 6": excerpt = "The first four characters of a UIC must be numbers."
        Case "COUNSELING": page = "1-13": pdfPage = 42: locator = "paragraph 1-2, DATE COUNSELED, block 30": excerpt = "If counseling was not performed for any other reason, enter ""NOT PERF."""
        Case "COUNSELOR": page = "1-13": pdfPage = 42: locator = "paragraph 1-2, COUNSELOR, block 31": excerpt = "If the entry in block 30 is NOT PERF, enter a brief explanation in block 31"
        Case "CAREER-PRESENT": page = "1-15": pdfPage = 44: locator = "paragraph 1-2, CAREER RECOMMENDATIONS": excerpt = "If no recommendation is appropriate, enter ""NA"" or ""NONE"" in the first block."
        Case "PARTIAL-NOB-LIMIT", "PARTIAL-NOB-MARKS": page = "6-1": pdfPage = 91: locator = "paragraph 6-3": excerpt = "Grade up to three traits and mark all other traits and the promotion recommendation as ""NOB."""
        Case "PARTIAL-NOB-ADVERSE": page = "6-1": pdfPage = 91: locator = "paragraph 6-3": excerpt = "An Observed report with an ""NOB"" promotion recommendation cannot be submitted if the member receives a 1.0 in any trait"
        Case "ADVERSE-REFERRAL": page = "I-2": pdfPage = 28: locator = "Introduction, paragraph 4a-c": excerpt = "containing adverse matter to be referred to the Service member for a statement before it is placed in the OMPF."
        Case "JUSTIFY-TRAITS": page = "13-3": pdfPage = 125: locator = "paragraph 13-4a-b": excerpt = "Specifically substantiate all 1.0 grades and when three or more traits are evaluated 2.0."
        Case "JUSTIFY-ADVERSE": page = "13-3": pdfPage = 125: locator = "paragraph 13-4c(1)": excerpt = "Substantiate any promotion recommendation of Significant Problems and any recommendation against retention."
        Case "REGULAR-LENGTH": page = "3-1": pdfPage = 70: locator = "paragraph 3-2": excerpt = "may not cover more than 15 months without PERS-32 approval."
        Case Else
            If Left$(code, 4) = "APP-" Then ValidationPolicyReference = "Application rule - no Navy policy citation applies." Else ValidationPolicyReference = "A source excerpt has not been verified for this rule. Refer to the original rule pack's locator."
            Exit Function
    End Select
    ValidationPolicyReference = "BUPERSINST 1610.10H CH-2 (26 May 2026 compilation)" & vbCrLf & "Enclosure (2), printed page " & page & "; " & locator & " (PDF page " & pdfPage & ")." & vbCrLf & "Excerpt: " & Chr$(34) & excerpt & Chr$(34) & vbCrLf & "Source SHA-256: " & NAVY_RULE_SOURCE_HASH
End Function
Private Sub VAdd(ByVal findings As Collection, ByVal f As Object, ByVal id As Long, ByVal key As String, ByVal severity As String, ByVal code As String, ByVal message As String)
    Dim item As Object
    If severity = "Policy issue" Then
        If Not IsDate(VText(f, "PeriodEnd")) Then
            severity = "Needs review": message = message & " Confirm the ending date and applicable instruction."
        ElseIf CDate(f("PeriodEnd")) < #5/26/2026# Or CDate(f("PeriodEnd")) > Date Then
            severity = "Needs review": message = message & " Verify the applicable instruction for this report date; this check uses the May 2026 compilation."
        End If
    End If
    Set item = CreateObject("Scripting.Dictionary")
    item.Add "ReportID", id: item.Add "Key", key: item.Add "Code", code: item.Add "Severity", severity
    item.Add "Member", UXIdentity(VText(f, "OfficialName"), VText(f, "MemberName"))
    item.Add "Label", UXFieldCaption(key, VText(f, "FormKind"))
    item.Add "Message", message: item.Add "Reference", ValidationPolicyReference(code): findings.Add item
End Sub
Public Function IsValidPolicyUIC(ByVal value As String) As Boolean
    If Len(value) <> 5 Then Exit Function
    IsValidPolicyUIC = (Left$(value, 4) Like "####" And UCase$(Right$(value, 1)) Like "[A-Z0-9]")
End Function
Private Sub VPolicyDetails(ByVal findings As Collection, ByVal f As Object, ByVal id As Long, ByVal observed As Boolean, ByVal graded As Long, ByVal twos As Long)
    Dim key As Variant, value As String, title As String, justify As Boolean, partial As Boolean, adverse As Boolean, referral As Boolean
    For Each key In Array(Array("ShipStation", "SHIP-PRESENT"), Array("PromotionStatus", "STATUS-PRESENT"), Array("DateReported", "REPORTED-PRESENT"), Array("ReportingSeniorName", "RS-NAME"), Array("ReportingSeniorGrade", "RS-GRADE"), Array("ReportingSeniorTitle", "RS-TITLE"), Array("ReportingSeniorUIC", "RS-UIC"))
        If VText(f, CStr(key(0))) = "" Then VAdd findings, f, id, CStr(key(0)), "Incomplete draft", CStr(key(1)), "Required report information is missing. Drafts may still be saved."
    Next
    For Each key In Array(Array("StatementYes", "StatementNo"), Array("RetentionYes", "RetentionNo"))
        If UXFieldApplies(CStr(key(0)), VText(f, "FormKind")) Then
            If VText(f, CStr(key(0))) = "Yes" And VText(f, CStr(key(1))) = "Yes" Then VAdd findings, f, id, CStr(key(0)), "Invalid entry", "APP-EXCLUSIVE", "Yes and No cannot both be selected for this response."
        End If
    Next
    For Each key In Array("UnitUIC", "ReportingSeniorUIC")
        value = VText(f, CStr(key))
        If value <> "" And Not IsValidPolicyUIC(value) Then VAdd findings, f, id, CStr(key), "Invalid entry", "UIC-FORMAT", "Use four digits followed by one letter or digit."
    Next
    value = VText(f, "CounselingDate")
    If value = "" Then VAdd findings, f, id, "CounselingDate", "Incomplete draft", "COUNSELING", "Enter the counseling date, NOT REQ when outside this period, or NOT PERF with an explanation."
    If IsValidIsoDate(value) Then
        If VText(f, "CounselorName") = "" Then VAdd findings, f, id, "CounselorName", "Incomplete draft", "COUNSELOR", "Enter the counselor from the counseling worksheet."
        If IsDate(VText(f, "PeriodStart")) And IsDate(VText(f, "PeriodEnd")) Then
            If CDate(value) < CDate(f("PeriodStart")) Or CDate(value) > CDate(f("PeriodEnd")) Then VAdd findings, f, id, "CounselingDate", "Policy issue", "COUNSELING", "This counseling date is outside the report period. Verify whether NOT REQ applies; retain the actual date in the appropriate period."
        End If
    End If
    If value = "NOT PERF" And VText(f, "CounselorName") = "" Then VAdd findings, f, id, "CounselorName", "Incomplete draft", "COUNSELOR", "Explain why counseling was not performed, or enter SEE COMMENT and explain in comments."
    If UCase$(VText(f, "CounselorName")) = "SEE COMMENT" And VText(f, "OfficialComments") = "" Then VAdd findings, f, id, "OfficialComments", "Incomplete draft", "COUNSELOR", "SEE COMMENT requires an explanation in the comments block."
    If observed And VText(f, "PromotionRecommendation") <> "NOB" And VText(f, "AssignmentRecommendationA") = "" Then VAdd findings, f, id, "AssignmentRecommendationA", "Incomplete draft", "CAREER-PRESENT", "Enter the first career recommendation, or NA / NONE if appropriate. The second recommendation is optional."
    partial = observed And VText(f, "PromotionRecommendation") = "NOB"
    If partial And graded > 3 Then VAdd findings, f, id, "PromotionRecommendation", "Policy issue", "PARTIAL-NOB-LIMIT", "An observed report with NOB promotion recommendation permits at most three graded traits."
    For Each key In TraitKeys()
        value = VText(f, CStr(key)): title = UCase$(CatalogueTraitTitle(VText(f, "FormKind"), CLng(Mid$(CStr(key), 6))))
        If partial And value = "" Then VAdd findings, f, id, CStr(key), "Incomplete draft", "PARTIAL-NOB-MARKS", "Mark ungraded traits NOB on a partially graded observed report."
        If value = "1.0" Then justify = True: adverse = True: referral = True
        If (value = "1.0" Or value = "2.0") And (InStr(title, "CLIMATE") > 0 Or InStr(title, "CHARACTER") > 0) Then justify = True: adverse = True
        If value = "1.0" Or value = "2.0" Then
            If VText(f, "FormKind") = "CHIEFEVAL" Then
                If InStr(title, "CHARACTER") > 0 Then referral = True
            Else
                If InStr(title, "CLIMATE") > 0 Then referral = True
            End If
        End If
    Next
    If twos >= 3 Then referral = True
    If referral Or VText(f, "PromotionRecommendation") = "Significant Problems" Then VAdd findings, f, id, "StatementYes", "Needs review", "ADVERSE-REFERRAL", "Verify referral to the member for a statement and any written decision not to submit one. Checkbox values and local approvals do not establish that referral occurred."
    If twos >= 3 Then justify = True: adverse = True
    If partial And adverse Then VAdd findings, f, id, "PromotionRecommendation", "Policy issue", "PARTIAL-NOB-ADVERSE", "NOB promotion recommendation conflicts with a 1.0 trait, three 2.0 traits, or climate/character below 3.0. Adverse narrative information also requires human review."
    If justify Then VAdd findings, f, id, "OfficialComments", "Needs review", "JUSTIFY-TRAITS", "Substantiate the adverse trait marks and correlate comments by block number. The app cannot establish whether the narrative adequately justifies the marks."
    If VText(f, "PromotionRecommendation") = "Significant Problems" Or VText(f, "RetentionNo") = "Yes" Then VAdd findings, f, id, "OfficialComments", "Needs review", "JUSTIFY-ADVERSE", "Substantiate Significant Problems / recommendation against retention and verify required adverse-report referral procedures."
    If VText(f, "ReportRegular") = "Yes" And IsDate(VText(f, "PeriodStart")) And IsDate(VText(f, "PeriodEnd")) Then
        If CDate(f("PeriodEnd")) > DateAdd("m", 15, CDate(f("PeriodStart"))) Then VAdd findings, f, id, "PeriodEnd", "Needs review", "REGULAR-LENGTH", "Regular report exceeds 15 months. Verify PERS-32 approval, including any letter extensions; approval is not recorded by this app."
    End If
End Sub
Public Function ValidateLiveFields(ByVal f As Object, Optional ByVal reportID As Long = 0) As Collection
    Dim findings As New Collection, key As Variant, text As String, options As String, allowed As Boolean, opt As Variant
    Dim count As Long, recommendation As String, observed As Boolean, graded As Long, kind As String, code As String, twos As Long, title As String
    kind = VText(f, "FormKind")
    For Each key In DraftFieldKeys()
        If Not UXFieldLegacy(CStr(key)) And UXFieldControl(CStr(key)) <> "" And UXFieldApplies(CStr(key), kind) Then
            text = VText(f, CStr(key))
            If Len(text) > 0 Then
                If UXFieldLimit(CStr(key)) > 0 And Len(text) > UXFieldLimit(CStr(key)) Then VAdd findings, f, reportID, CStr(key), "Invalid entry", "APP-LENGTH", "Use " & UXFieldLimit(CStr(key)) & " characters or fewer."
                If UXFieldDigits(CStr(key)) > 0 And key <> "UnitUIC" And key <> "ReportingSeniorUIC" Then
                    If Len(text) <> UXFieldDigits(CStr(key)) Or text Like "*[!0-9]*" Then VAdd findings, f, reportID, CStr(key), "Invalid entry", "APP-DIGITS", "Enter exactly " & UXFieldDigits(CStr(key)) & " digits (stored without punctuation)."
                End If
                If UXFieldDate(CStr(key)) And Not IsDate(text) And Not (key = "CounselingDate" And IsCounselingStatus(text)) Then VAdd findings, f, reportID, CStr(key), "Invalid entry", "APP-DATE", "Enter a valid date."
                options = UXFieldOptions(CStr(key)): allowed = False
                If Len(options) > 0 Then
                    For Each opt In Split(options, ";"): If text = CStr(opt) Then allowed = True
                    Next
                    If Not allowed Then VAdd findings, f, reportID, CStr(key), "Invalid entry", "APP-OPTION", "Choose a listed value."
                End If
            End If
        End If
    Next
    For Each key In Array("PeriodStart", "PeriodEnd", "ReportingSeniorSSN")
        If VText(f, CStr(key)) = "" Then VAdd findings, f, reportID, CStr(key), "Incomplete draft", IIf(key = "ReportingSeniorSSN", "RS-SSN", "PERIOD"), IIf(key = "ReportingSeniorSSN", "Reporting senior SSN is needed for OMPF filing; it may be omitted from some member copies.", "Enter the report period. Drafts may still be saved.")
    Next
    For Each key In Array("OfficialName", "RankRate", "UnitUIC", "MemberSSN")
        If VText(f, CStr(key)) = "" Then
            Select Case key
                Case "OfficialName": code = "NAME-PRESENT"
                Case "RankRate": code = "RATE-PRESENT"
                Case "UnitUIC": code = "UIC-PRESENT"
                Case "MemberSSN": code = "SSN-PRESENT"
            End Select
            VAdd findings, f, reportID, CStr(key), "Incomplete draft", code, "Required report information is missing. Drafts may still be saved."
        End If
    Next
    If IsDate(VText(f, "PeriodStart")) And IsDate(VText(f, "PeriodEnd")) Then
        If CDate(f("PeriodStart")) > CDate(f("PeriodEnd")) Then VAdd findings, f, reportID, "PeriodStart", "Invalid entry", "APP-PERIOD", "Start date is after end date."
    End If
    count = 0
    For Each key In Array("ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW")
        If VText(f, CStr(key)) = "Yes" Then count = count + 1
    Next
    If count <> 1 Then VAdd findings, f, reportID, "ComponentActive", IIf(count = 0, "Incomplete draft", "Policy issue"), "COMPONENT", "Select exactly one duty/competitive status."
    count = 0
    For Each key In Array("ReportRegular", "ReportConcurrent", "ReportOperationalCommander")
        If VText(f, CStr(key)) = "Yes" Then count = count + 1
    Next
    If count = 0 Then VAdd findings, f, reportID, "ReportRegular", "Incomplete draft", "REPORT-TYPE", "Select a report type."
    If count > 1 And VText(f, "ReportOperationalCommander") = "Yes" Then VAdd findings, f, reportID, "ReportOperationalCommander", "Needs review", "REPORT-TYPE", "Regular + Concurrent is supported; verify other combined types."
    count = 0
    For Each key In Array("OccasionPeriodic", "OccasionDetachment", "OccasionFrocking", "OccasionSpecial")
        If VText(f, CStr(key)) = "Yes" Then count = count + 1
    Next
    If count = 0 Then VAdd findings, f, reportID, "OccasionPeriodic", "Incomplete draft", "OCCASION", "Select at least one applicable occasion."
    If count > 1 And VText(f, "OccasionSpecial") = "Yes" Then VAdd findings, f, reportID, "OccasionSpecial", "Policy issue", "OCCASION", "Special cannot be combined with another report occasion."
    recommendation = VText(f, "PromotionRecommendation"): observed = (VText(f, "ObservationStatus") <> "NOB")
    If recommendation = "" Then VAdd findings, f, reportID, "PromotionRecommendation", "Incomplete draft", "PROMOTION-PRESENT", "Choose a promotion recommendation appropriate to this report."
    For Each key In TraitKeys()
        text = VText(f, CStr(key))
        If text <> "" And text <> "NOB" Then graded = graded + 1
        If text = "2.0" Then twos = twos + 1
        If observed And recommendation <> "NOB" And text = "" Then VAdd findings, f, reportID, CStr(key), "Incomplete draft", "APP-TRAIT", "Choose a trait mark or NOB as appropriate before final review."
        If Not observed And text <> "" Then VAdd findings, f, reportID, CStr(key), "Policy issue", "WHOLE-NOB", "Wholly NOB reports leave trait grades blank; review chapters 6/10 exceptions if applicable."
        If observed And text = "1.0" And (recommendation = "Promotable" Or recommendation = "Must Promote" Or recommendation = "Early Promote") Then VAdd findings, f, reportID, CStr(key), "Policy issue", "TRAIT-RECOMMENDATION", "A 1.0 trait conflicts with a Promotable or higher recommendation."
        If observed And text = "2.0" And (recommendation = "Must Promote" Or recommendation = "Early Promote") Then VAdd findings, f, reportID, CStr(key), "Policy issue", "TRAIT-TWO", "A 2.0 trait conflicts with Must Promote or Early Promote."
        If observed And text = "2.0" And recommendation = "Promotable" Then
            title = UCase$(CatalogueTraitTitle(kind, CLng(Mid$(CStr(key), 6))))
            If title = "ACCOUNTABILITY" Or InStr(title, "CLIMATE") > 0 Then VAdd findings, f, reportID, CStr(key), "Policy issue", "TRAIT-CRITICAL", "This trait must be at least 3.0 for a Promotable or higher recommendation."
            If InStr(title, "CHARACTER") > 0 Then VAdd findings, f, reportID, CStr(key), "Needs review", "TRAIT-TWO-COUNT", "Character is excluded from the ordinary allowance for two 2.0 traits; verify advancement eligibility and recommendation."
        End If
    Next
    If observed And twos > 2 And recommendation = "Promotable" Then VAdd findings, f, reportID, "PromotionRecommendation", "Policy issue", "TRAIT-TWO-COUNT", "Promotable permits at most two qualifying 2.0 traits; this report has " & twos & "."
    If Not observed Then
        If recommendation <> "NOB" Then VAdd findings, f, reportID, "PromotionRecommendation", "Policy issue", "WHOLE-NOB", "A wholly NOB report requires an NOB promotion recommendation."
        For Each key In Array("AssignmentRecommendationA", "AssignmentRecommendationB")
            If UXFieldApplies(CStr(key), kind) And VText(f, CStr(key)) <> "" Then VAdd findings, f, reportID, CStr(key), "Policy issue", "WHOLE-NOB", "Wholly NOB reports leave career recommendations blank."
        Next
    ElseIf recommendation = "NOB" Then
        VAdd findings, f, reportID, "PromotionRecommendation", "Needs review", "PARTIAL-NOB", "Partially graded NOB: verify short-period applicability, at most three graded traits and narrative justification. Current graded traits: " & graded & "."
    End If
    If kind = "FITREP" And (VText(f, "RankRate") = "ENS" Or VText(f, "RankRate") = "LTJG") And (recommendation = "Must Promote" Or recommendation = "Early Promote") Then
        If Len(VText(f, "Designator")) = 4 And Not VText(f, "Designator") Like "*[!0-9]*" Then
            If Left$(VText(f, "Designator"), 1) <> "6" Then VAdd findings, f, reportID, "PromotionRecommendation", "Policy issue", "JUNIOR-OFFICER", "Non-LDO ENS/LTJG reports cannot exceed Promotable."
        Else
            VAdd findings, f, reportID, "Designator", "Needs review", "JUNIOR-OFFICER", "A valid officer designator is needed to evaluate the LDO exception."
        End If
    End If
    VPolicyDetails findings, f, reportID, observed, graded, twos
    Set ValidateLiveFields = findings
End Function
Private Function VComponent(ByVal f As Object) As String
    Dim key As Variant
    For Each key In Array("ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW")
        If VText(f, CStr(key)) = "Yes" Then
            If VText(f, "FormKind") <> "FITREP" And (key = "ComponentActive" Or key = "ComponentTAR") Then VComponent = VComponent & "ACT/TAR" Else VComponent = VComponent & CStr(key)
        End If
    Next
End Function
Public Function ValidationWornPaygrade(ByVal rate As String) As String
    rate = UCase$(Trim$(rate))
    Select Case rate
        Case "ENS": ValidationWornPaygrade = "O1"
        Case "LTJG": ValidationWornPaygrade = "O2"
        Case "LT": ValidationWornPaygrade = "O3"
        Case "LCDR": ValidationWornPaygrade = "O4"
        Case "CDR": ValidationWornPaygrade = "O5"
        Case "CAPT": ValidationWornPaygrade = "O6"
        Case "WO1": ValidationWornPaygrade = "W1"
        Case "CWO2", "CWO3", "CWO4", "CWO5": ValidationWornPaygrade = "W" & Right$(rate, 1)
        Case Else
            If Right$(rate, 2) = "CM" Then ValidationWornPaygrade = "E9": Exit Function
            If Right$(rate, 2) = "CS" Then ValidationWornPaygrade = "E8": Exit Function
            If Right$(rate, 1) = "C" Then ValidationWornPaygrade = "E7": Exit Function
            Select Case Right$(rate, 1)
                Case "1": ValidationWornPaygrade = "E6"
                Case "2": ValidationWornPaygrade = "E5"
                Case "3": ValidationWornPaygrade = "E4"
            End Select
            Select Case Right$(rate, 2)
                Case "SR", "FR", "AR", "CR", "HR": ValidationWornPaygrade = "E1"
                Case "SA", "FA", "AA", "CA", "HA": ValidationWornPaygrade = "E2"
                Case "SN", "FN", "AN", "CN", "HN": ValidationWornPaygrade = "E3"
            End Select
    End Select
End Function
Public Function VCategory(ByVal f As Object, Optional ByVal officerBasis As String = "Unconfirmed") As String
    Dim designator As String, prefix As String
    designator = VText(f, "Designator")
    If Len(designator) <> 4 Or designator Like "*[!0-9]*" Then Exit Function
    prefix = Left$(designator, 2)
    Select Case prefix
        Case "11", "13", "19": VCategory = "URL"
        Case "12"
            If Left$(designator, 3) = "123" Or Left$(designator, 3) = "128" Then VCategory = Left$(designator, 3) Else VCategory = "HR"
        Case "14", "17": VCategory = prefix
        Case Else
            Select Case Left$(designator, 3)
                Case "150", "151", "152", "154", "165", "166", "168", "180", "181", "182", "183", "184", "186", "187", "188", "210", "220", "230", "250", "270", "290", "310", "410", "510": VCategory = Left$(designator, 3)
            End Select
    End Select
    'Duty status is not commission/category basis. Require a deliberate group review.
    If officerBasis = "Reserve" Then
        Select Case prefix
            Case "61", "62", "63", "64": VCategory = "Reserve LDO line"
            Case "65": VCategory = "Reserve LDO staff"
        End Select
        If Left$(designator, 1) = "7" Then VCategory = "Reserve CWO"
    ElseIf officerBasis = "Active" Then
        Select Case prefix
            Case "61", "62", "63", "64", "68": VCategory = "Active LDO " & prefix
            Case "71", "72", "73", "74", "75", "78": VCategory = "Active CWO " & prefix
        End Select
        If Left$(designator, 3) = "651" Or Left$(designator, 3) = "653" Then VCategory = "Active LDO " & Left$(designator, 3)
        If Left$(designator, 3) = "740" Then VCategory = "Active CWO 72"
        If designator = "7371" Then VCategory = "Active CWO UAV"
    End If
End Function
Public Function GroupObservedRecommendation(ByVal f As Object) As Boolean
    If VText(f, "ObservationStatus") = "NOB" Then Exit Function
    Select Case VText(f, "PromotionRecommendation")
        Case "Significant Problems", "Progressing", "Promotable", "Must Promote", "Early Promote": GroupObservedRecommendation = True
    End Select
End Function
Public Function GroupKeysResolved(ByVal f As Object, ByVal grade As String, ByVal kind As String, Optional ByVal officerBasis As String = "Unconfirmed") As Boolean
    Dim key As Variant, count As Long, ssn As String
    If VText(f, "FormKind") <> kind Or ValidationWornPaygrade(VText(f, "RankRate")) <> grade Then Exit Function
    If Not IsDate(VText(f, "PeriodEnd")) Or VText(f, "PromotionStatus") = "" Then Exit Function
    For Each key In Array("ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW")
        If VText(f, CStr(key)) = "Yes" Then count = count + 1
    Next
    If count <> 1 Then Exit Function
    If VText(f, "ReportRegular") <> "Yes" And VText(f, "ReportConcurrent") <> "Yes" And VText(f, "ReportOperationalCommander") <> "Yes" Then Exit Function
    ssn = Replace(VText(f, "ReportingSeniorSSN"), "-", "")
    If Len(ssn) <> 9 Or ssn Like "*[!0-9]*" Or VText(f, "ReportingSeniorName") = "" Then Exit Function
    If kind = "FITREP" And VCategory(f, officerBasis) = "" Then Exit Function
    If kind <> "FITREP" And VText(f, "UnitUIC") = "" Then Exit Function
    GroupKeysResolved = True
End Function
Public Function ValidateLiveGroup(ByVal groupID As Long, Optional ByVal overrides As Object, Optional ByVal membersCSV As String = "@Current") As Collection
    Dim findings As New Collection, rows As New Collection, r As DAO.Recordset, db As DAO.Database, f As Object, item As Object, row As Object, seen As Object, ids As Object, csv As String, officerBasis As String, uicMode As String
    Dim key As Variant, value As String, id As Long, code As String, ep As Long, mp As Long, observed As Long, capacity As Object, grade As String, ldo As Boolean, unknownLDO As Boolean, nonLDO As Boolean, keysUnresolved As Boolean, keysMixed As Boolean, worn As String, expectedKind As String, severity As String
    grade = Nz(DLookup("Paygrade", "SummaryGroup", "GroupID=" & groupID), "")
    expectedKind = Nz(DLookup("FormKind", "SummaryGroup", "GroupID=" & groupID), "")
    officerBasis = GroupPolicyValue(groupID, "OfficerBasis"): uicMode = GroupPolicyValue(groupID, "UICMode")
    Set db = CurrentDb
    If membersCSV = "@Current" Then
        Set r = db.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & groupID & " ORDER BY ReportID", dbOpenSnapshot)
    Else
        csv = "0"
        If Len(membersCSV) > 0 Then
            Set ids = MajorIDs(membersCSV)
            For Each key In ids.Keys: csv = csv & "," & CLng(ids(key)): Next
        End If
        Set r = db.OpenRecordset("SELECT ReportID FROM Report WHERE ReportID In (" & csv & ") ORDER BY ReportID", dbOpenSnapshot)
    End If
    Do While Not r.EOF
        id = r!ReportID: Set f = SavedFields(id)
        If Not overrides Is Nothing Then
            If overrides.Exists(CStr(id)) Then Set f = overrides(CStr(id))
        ElseIf membersCSV = "@Current" And CurrentProject.AllForms("frmEditor").IsLoaded Then
            If Nz(Forms!frmEditor!txtReportID, 0) = id Then Set f = EditorPendingFields()
        End If
        Set row = CreateObject("Scripting.Dictionary"): row.Add "ID", id: row.Add "Fields", f: rows.Add row
        For Each item In ValidateLiveFields(f, id): findings.Add item: Next
        If VText(f, "FormKind") = "FITREP" Then code = "GROUP-OFFICER" Else code = "GROUP-ENLISTED"
        If VText(f, "FormKind") <> expectedKind Then VAdd findings, f, id, "FormKind", "Policy issue", code, "Report type differs from the summary group's " & expectedKind & " type."
        worn = ValidationWornPaygrade(VText(f, "RankRate"))
        If worn = "" Then
            VAdd findings, f, id, "RankRate", "Needs review", code, "Verify worn paygrade; this grade/rate could not be resolved automatically. Group paygrade: " & grade & "."
        ElseIf worn <> grade Then
            VAdd findings, f, id, "RankRate", "Policy issue", code, "Worn paygrade " & worn & " differs from group paygrade " & grade & ". Ratings may differ, but current paygrade must match."
        End If
        If VText(f, "FormKind") = "FITREP" And VCategory(f, officerBasis) = "" Then VAdd findings, f, id, "Designator", "Needs review", "GROUP-CATEGORY", "Verify officer competitive category, including active/reserve LDO and CWO distinctions. This category is outside the automatic mapping."
        If VText(f, "PromotionRecommendation") = "NOB" Or VText(f, "ObservationStatus") = "NOB" Then VAdd findings, f, id, "PromotionRecommendation", "Policy issue", code, "NOB promotion recommendations are excluded from observed summary groups."
        If GroupObservedRecommendation(f) Then
            observed = observed + 1
            If uicMode = "Separate enlisted UICs" Or uicMode = "Officer dual-hatted" Then
                If VText(f, "UnitUIC") = "" Or VText(f, "ReportingSeniorUIC") = "" Or VText(f, "UnitUIC") <> VText(f, "ReportingSeniorUIC") Then
                    keysUnresolved = True
                    VAdd findings, f, id, "UnitUIC", "Policy issue", code, "This confirmed UIC breakout requires member UIC (block 6) to match reporting senior UIC (block 26). Verify the recorded grouping authority."
                End If
            End If
            If VText(f, "PromotionRecommendation") = "Early Promote" Then ep = ep + 1
            If VText(f, "PromotionRecommendation") = "Must Promote" Then mp = mp + 1
            If Not GroupKeysResolved(f, grade, expectedKind, officerBasis) Then
                keysUnresolved = True
                VAdd findings, f, id, "RankRate", "Needs review", code, "Observed roster keys are unresolved. Confirm worn paygrade, promotion status, ending date, one duty status, report type, reporting senior identity and applicable UIC/competitive category before calculating capacity."
            End If
        ElseIf VText(f, "ObservationStatus") <> "NOB" And VText(f, "PromotionRecommendation") <> "NOB" Then
            keysUnresolved = True
        End If
        If GroupObservedRecommendation(f) And VText(f, "FormKind") = "FITREP" Then
            If Len(VText(f, "Designator")) <> 4 Or VText(f, "Designator") Like "*[!0-9]*" Then
                unknownLDO = True
            ElseIf Left$(VText(f, "Designator"), 1) = "6" Then
                ldo = True
            Else
                nonLDO = True
            End If
        End If
        r.MoveNext
    Loop
    r.Close
    For Each key In Array("FormKind", "PromotionStatus", "PeriodEnd", "ComponentActive", "ReportRegular", "ReportConcurrent", "ReportOperationalCommander", "BilletSubcategory", "ReportingSeniorName", "ReportingSeniorSSN", "UnitUIC", "ReportingSeniorUIC", "Designator")
        Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
        For Each row In rows
            Set f = row("Fields"): value = VText(f, CStr(key))
            If key = "ComponentActive" Then value = VComponent(f)
            If key = "PeriodEnd" And IsDate(value) Then value = Format$(CDate(value), "yyyy-mm-dd")
            If key = "Designator" Then
                If VText(f, "FormKind") <> "FITREP" Then value = "Enlisted" Else value = VCategory(f, officerBasis)
            End If
            If key = "UnitUIC" And VText(f, "FormKind") = "FITREP" Then value = "Officer UIC: contextual review"
            If key = "ReportingSeniorUIC" And uicMode <> "Officer dual-hatted" And uicMode <> "Separate enlisted UICs" Then value = "Not an explicit UIC breakout"
            If key = "UnitUIC" And uicMode = "Combined enlisted UICs" And VText(f, "FormKind") <> "FITREP" Then value = "Combined enlisted UICs"
            If GroupObservedRecommendation(f) Then seen(value) = True
        Next
        If seen.count > 1 Then
            If key <> "ReportingSeniorName" Then keysMixed = True
            severity = "Policy issue"
            If seen.Exists("") Or key = "UnitUIC" Or key = "Designator" Or key = "ReportingSeniorName" Then severity = "Needs review"
            For Each row In rows
                Set f = row("Fields")
                If VText(f, "FormKind") = "FITREP" Then code = "GROUP-OFFICER" Else code = "GROUP-ENLISTED"
                If key = "Designator" Then code = "GROUP-CATEGORY"
                VAdd findings, f, CLng(row("ID")), CStr(key), severity, code, "Group members differ in " & UXFieldCaption(CStr(key), VText(f, "FormKind")) & ". This report: " & VText(f, CStr(key)) & ". Verify grouping; UIC exceptions, officer categories and name variations require judgment."
            Next
        End If
    Next
    If rows.count > 0 Then
        Set row = rows(1): Set f = row("Fields"): grade = Nz(DLookup("Paygrade", "SummaryGroup", "GroupID=" & groupID), "")
        If VText(f, "FormKind") = "FITREP" Then code = "GROUP-OFFICER" Else code = "GROUP-ENLISTED"
        VAdd findings, f, CLng(row("ID")), "RankRate", "Needs review", code, GroupReviewLabel(groupID)
        If observed > 0 And Not keysUnresolved And Not keysMixed And Not ((grade = "O1" Or grade = "O2") And (unknownLDO Or (ldo And nonLDO))) Then
            On Error Resume Next
            Set capacity = PromotionCapacity(grade, observed, ldo, 0)
            On Error GoTo 0
            If Not capacity Is Nothing Then
                value = "Provisional local roster: " & observed & " observed; EP " & ep & "/" & capacity("EarlyPromoteMaximum") & "; EP + MP " & (ep + mp) & "/" & capacity("CombinedMaximum") & ". Verify complete eligible membership and group keys before relying on these limits."
                severity = "Needs review"
                If ep > capacity("EarlyPromoteMaximum") Or ep + mp > capacity("CombinedMaximum") Then
                    value = "Local observed roster exceeds the provisional limit. " & value: severity = "Policy issue"
                End If
                VAdd findings, f, CLng(row("ID")), "PromotionRecommendation", severity, "CAPACITY", value
            End If
        End If
        If Not overrides Is Nothing Or membersCSV <> "@Current" Then VAdd findings, f, CLng(row("ID")), "RankRate", "Needs review", code, "Proposed changes are outside the recorded saved-roster review. Reconfirm completeness after Apply."
        If capacity Is Nothing Then VAdd findings, f, CLng(row("ID")), "PromotionRecommendation", "Needs review", "CAPACITY", "Capacity not calculated: the observed roster is empty, incomplete, mixed or outside supported competitive categories. Resolve group findings before evaluating EP/MP limits. NOB reports are excluded from both counts and limits; complete command membership still requires verification."
    End If
    Set ValidateLiveGroup = findings
End Function
'Live checks are bounded diagnostics, not Navy-ready certification. No database mutations.
'Historical rule definitions remain immutable; storage rules have no invented policy source.
