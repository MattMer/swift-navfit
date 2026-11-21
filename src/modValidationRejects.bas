Attribute VB_Name = "modValidationRejects"
Option Compare Binary
Option Explicit
' Checks drawn from the reasons PERS-32 gives for rejecting reports (its published
' numerical error codes) and the EVALMAN paragraphs behind them. Each check cites the
' built-in EVALMAN (modEvalmanText) by section, with an excerpt that the source
' contract test finds word for word in src/evalman.json. A reason that appears only on
' the PERS-32 page cites that page and no excerpt. Signatures, which cause many
' rejections, are on paper and cannot be checked here.

Public Const PERS32_REJECT_SOURCE As String = "NAVPERSCOM PERS-32, Performance Evaluation Error Codes & Reasons (mynavyhr.navy.mil), read 20 Nov 2026"

Private Function RText(ByVal f As Object, ByVal key As String) As String
    If f.Exists(key) Then RText = Trim$(Nz(f(key), ""))
End Function

' Whole-word match, so "CO" is not found inside "COMMAND".
Private Function RHasWord(ByVal text As String, ByVal word As String) As Boolean
    Dim c As Variant
    text = " " & UCase$(text) & " "
    For Each c In Array("/", ",", ".", "-", "(", ")", ";", ":")
        text = Replace(text, CStr(c), " ")
    Next
    RHasWord = (InStr(text, " " & UCase$(word) & " ") > 0)
End Function

' The EVALMAN section (start of its title), its excerpt, and the PERS-32 reject codes.
Private Function RejectRule(ByVal code As String, ByVal part As Long) As String
    Dim t As String, e As String, c As String
    Select Case code
        Case "RJ-SUBCAT-BLANK": t = "Block 21:": c = "21B": e = "Select or enter the billet subcategory code, if authorized, or enter ""NA."" Do not leave blank."
        Case "RJ-SUBCAT-DUTY": t = "Block 21:": c = "21D": e = "If annotated with an * below, Block 21 should match entry in Block 29."
        Case "RJ-FROCK-STATUS": t = "3-8.": c = "NOTE 2/I": e = "Enter the old rate in block 2, enter ""Selected"" in block 8, and show the ending date as the date prior to frocking (or promotion if not previously frocked)."
        Case "RJ-FROCK-OFFICER": t = "3-8.": c = "10-13G": e = "Do not submit a promotion/frocking report for an officer or enlisted member who has been promoted, unless the change will result in more than 15 months between Regular reports"
        Case "RJ-ENS-SELECTED": t = "3-8.": c = "8A": e = "For ensigns promoting to LTJG, enter ""REGULAR"" in block 8, mark the report SPECIAL, and enter ""PROMOTING TO LTJG"" at the top of block 41."
        Case "RJ-SPECIAL": t = "3-9.": c = "10-13C, 10-13D": e = "State the justification for the Special report as the opening sentence of the comments"
        Case "RJ-OPS-CDR-MEMBER": t = "5-1.": c = "17-19C": e = "Operational Commander reports provide a supplemental record of performance for active-duty COs or OICs, as observed by their operational commanders."
        Case "RJ-OPS-CDR-TITLE": t = "5-2.": c = "17-19B": e = "An operational commander reporting senior must be a commander, other than the regular reporting senior"
        Case "RJ-ADVERSE-SIGN": t = "14-4.": c = "46B, 49B, 51B": e = "and the report is not adverse, the reporting senior should select the alternate signature or enter ""Certified Copy Provided"" in the member's signature block"
        Case "RJ-PFA-FAILURES": t = "Block 20:": c = "43L, 46L": e = "For reporting period in which a member accumulated three PFA failures in a 4-year period:"
        Case "RJ-PROGRESSING": t = "16-2.": c = "41D, 45D": e = """Progressing"" may not be marked if a ""Promotable"" or higher advancement recommendation is already in effect in the current grade, even if the recommendation came from a previous command."
        Case "RJ-DATE-REPORTED": t = "Block 9:": c = "NOTE 2/J"
        Case "RJ-FUTURE-END": t = "Blocks 14-15:": c = "14-15E"
        Case "RJ-DUPLICATE": t = "3-2.": c = "NOTE 2/C"
    End Select
    Select Case part
        Case 0: RejectRule = t
        Case 1: RejectRule = e
        Case 2: RejectRule = c
    End Select
End Function

' The EVALMAN section a reject finding opens.
Public Function RejectRuleSection(ByVal code As String) As String
    RejectRuleSection = RejectRule(code, 0)
End Function

' The excerpt a reject finding quotes; "" when only the PERS-32 page states the reason.
Public Function RejectRuleExcerpt(ByVal code As String) As String
    RejectRuleExcerpt = RejectRule(code, 1)
End Function

Public Function RejectPolicyReference(ByVal code As String) As String
    Dim text As String, i As Long
    text = "PERS-32 reject reason " & RejectRule(code, 2) & ". " & PERS32_REJECT_SOURCE & "."
    i = EvalmanSectionByTitle(RejectRule(code, 0))
    If Len(RejectRule(code, 1)) > 0 And i >= 0 Then
        text = text & vbCrLf & EvalmanSource() & ", " & EvalmanHeader(i, 1) & " (printed page " & EvalmanHeader(i, 2) & ", PDF page " & EvalmanHeader(i, 3) & ")." & vbCrLf & "Excerpt: " & Chr$(34) & RejectRule(code, 1) & Chr$(34)
    Else
        text = text & vbCrLf & "EVALMAN has no sentence for this reason; the PERS-32 page is the source."
    End If
    RejectPolicyReference = text
End Function

' Checks for one report. Live, saved and group validation all call this.
Public Sub RejectChecks(ByVal findings As Collection, ByVal f As Object, ByVal id As Long)
    Dim kind As String, worn As String, recommendation As String, subcat As String, duty As String, key As Variant, title As String, mark As String
    Dim adverse As Boolean, twos As Long, bearing As String, observed As Boolean, starred As Variant, text As String
    kind = RText(f, "FormKind"): worn = ValidationWornPaygrade(RText(f, "RankRate"))
    recommendation = RText(f, "PromotionRecommendation"): observed = (RText(f, "ObservationStatus") <> "NOB")
    subcat = UCase$(RText(f, "BilletSubcategory")): duty = RText(f, "PrimaryDuty") & " " & RText(f, "DutyDescription")

    'Block 21 is never blank; NA is the usual entry.
    If subcat = "" Then
        VAdd findings, f, id, "BilletSubcategory", "Incomplete draft", "RJ-SUBCAT-BLANK", "Enter the billet subcategory, or NA when none applies. PERS-32 rejects a blank block 21."
    Else
        Select Case subcat
            Case "CO AFLOAT", "CO ASHORE"
                If Not (RHasWord(duty, "CO") Or InStr(UCase$(duty), "COMMANDING OFFICER") > 0) Then VAdd findings, f, id, "PrimaryDuty", "Needs review", "RJ-SUBCAT-DUTY", "Block 21 is " & subcat & ", but block 29 does not show the member as CO. PERS-32 rejects a billet subcategory the primary duty does not support."
            Case "OIC"
                If Not (RHasWord(duty, "OIC") Or InStr(UCase$(duty), "OFFICER IN CHARGE") > 0) Then VAdd findings, f, id, "PrimaryDuty", "Needs review", "RJ-SUBCAT-DUTY", "Block 21 is OIC, but block 29 does not show the member as officer in charge. PERS-32 rejects a billet subcategory the primary duty does not support."
            Case Else
                For Each starred In Array("CRF", "CANVASSER", "RESIDENT", "INTERN", "STUDENT")
                    If subcat = starred And InStr(UCase$(duty), CStr(starred)) = 0 Then VAdd findings, f, id, "PrimaryDuty", "Needs review", "RJ-SUBCAT-DUTY", "Block 21 is " & subcat & "; enter the same in block 29 (primary duty)."
                Next
        End Select
    End If

    'Adverse (EVALMAN introduction, paragraph 4a-b): a 1.0, Significant Problems, more
    'than two 2.0s, or below 3.0 in Equal Opportunity (E-1 to E-6, W-1 to O-6) or in the
    'E-7 to E-9 trait EVALMAN calls Character, which is block 37, Accountability, on the
    'current CHIEFEVAL (PERS-32 code 43G names block 37).
    For Each key In TraitKeys()
        mark = RText(f, CStr(key)): title = UCase$(CatalogueTraitTitle(kind, CLng(Mid$(CStr(key), 6))))
        If mark = "1.0" Then adverse = True
        If mark = "2.0" Then twos = twos + 1
        If mark = "1.0" Or mark = "2.0" Then
            If InStr(title, "CLIMATE") > 0 Or (kind = "CHIEFEVAL" And (InStr(title, "CHARACTER") > 0 Or InStr(title, "ACCOUNTABILITY") > 0)) Then adverse = True
        End If
        If InStr(title, "BEARING") > 0 Or (kind = "CHIEFEVAL" And InStr(title, "PROFESSIONALISM") > 0) Then bearing = mark
    Next
    If twos > 2 Or recommendation = "Significant Problems" Then adverse = True
    If adverse Then VAdd findings, f, id, "OfficialComments", "Needs review", "RJ-ADVERSE-SIGN", "This report is adverse: the member must sign it. ""Certified Copy Provided"" is not allowed; if the member refuses, the reporting senior notes that in the member's signature block and initials it."

    'Promotion/frocking reports: block 8 SELECTED, and only when they are required.
    If RText(f, "OccasionFrocking") = "Yes" Then
        If UCase$(RText(f, "PromotionStatus")) <> "SELECTED" Then VAdd findings, f, id, "PromotionStatus", "Policy issue", "RJ-FROCK-STATUS", "A Promotion/Frocking report shows SELECTED in block 8, the grade or rate before frocking in block 2 and the day before frocking in block 15."
        If kind <> "EVAL" Or worn <> "E6" Then VAdd findings, f, id, "OccasionFrocking", "Needs review", "RJ-FROCK-OFFICER", "A Promotion/Frocking report is required on frocking to E-7. Otherwise submit one only when the change would leave more than 15 months between Regular reports."
    End If
    'An ensign promoting to LTJG stays REGULAR in block 8 on a Special report.
    If kind = "FITREP" And worn = "O1" And UCase$(RText(f, "PromotionStatus")) = "SELECTED" Then VAdd findings, f, id, "PromotionStatus", "Policy issue", "RJ-ENS-SELECTED", "SELECTED is not allowed in block 8 for an ensign. For an ensign promoting to LTJG, enter REGULAR, mark the report Special and begin the comments with PROMOTING TO LTJG."

    'Special reports state why in the first sentence of the comments.
    If RText(f, "OccasionSpecial") = "Yes" Then
        text = "Begin the comments with why this Special report is submitted (for example, the promotion board it is for)."
        If kind = "FITREP" Then text = text & " Special reports for superior performance are not allowed for officers."
        If RText(f, "OfficialComments") = "" Then
            VAdd findings, f, id, "OfficialComments", "Incomplete draft", "RJ-SPECIAL", text
        Else
            VAdd findings, f, id, "OfficialComments", "Needs review", "RJ-SPECIAL", "Check that the first sentence of the comments says why this Special report is submitted." & IIf(kind = "FITREP", " Special reports for superior performance are not allowed for officers.", "")
        End If
    End If

    'Operational Commander reports: on a CO or OIC, by a commander.
    If RText(f, "ReportOperationalCommander") = "Yes" Then
        text = UCase$(RText(f, "PrimaryDuty") & " " & RText(f, "CommandEmployment") & " " & RText(f, "DutyDescription"))
        If Not (subcat = "CO AFLOAT" Or subcat = "CO ASHORE" Or subcat = "OIC" Or RHasWord(text, "CO") Or RHasWord(text, "OIC") Or InStr(text, "COMMANDING OFFICER") > 0 Or InStr(text, "OFFICER IN CHARGE") > 0) Then VAdd findings, f, id, "PrimaryDuty", "Needs review", "RJ-OPS-CDR-MEMBER", "An Operational Commander report is only for a CO or OIC. Show the member as CO or OIC in block 21, 28 or 29."
        text = UCase$(RText(f, "ReportingSeniorTitle"))
        If text <> "" And Not (InStr(text, "COM") > 0 Or InStr(text, "CDR") > 0 Or RHasWord(text, "CO") Or RHasWord(text, "OIC")) Then VAdd findings, f, id, "ReportingSeniorTitle", "Needs review", "RJ-OPS-CDR-TITLE", "On an Operational Commander report, block 25 shows the reporting senior as a commander or officer in charge."
    End If

    'Physical readiness: a third failure in 4 years sets the marks.
    If kind <> "CHIEFEVAL" And InStr(UCase$(RText(f, "ReadinessSummary")), "F") > 0 And observed Then
        If bearing <> "1.0" And recommendation <> "Significant Problems" Then VAdd findings, f, id, "ReadinessSummary", "Needs review", "RJ-PFA-FAILURES", "Block 20 shows a PFA failure. If it is the member's third in 4 years, Military Bearing must be 1.0 or the report marked Significant Problems" & IIf(kind = "EVAL", ", with retention Not Recommended", "") & ". Swift cannot see failures before this report."
    End If

    'Dates: reported before the end of the period, and the period already over.
    If IsDate(RText(f, "DateReported")) And IsDate(RText(f, "PeriodEnd")) Then
        If CDate(RText(f, "DateReported")) > CDate(RText(f, "PeriodEnd")) Then VAdd findings, f, id, "DateReported", "Policy issue", "RJ-DATE-REPORTED", "Date reported is after the end of the report period. The member was not at the command on the ending date, so this report is not authorized."
    End If
    If IsDate(RText(f, "PeriodEnd")) Then
        If CDate(RText(f, "PeriodEnd")) > Date Then VAdd findings, f, id, "PeriodEnd", "Information", "RJ-FUTURE-END", "The report period ends in the future. PERS-32 rejects a report submitted before its ending date."
    End If

    RejectHistoryChecks findings, f, id, kind, worn, recommendation
End Sub

' From this member's other reports in the workspace: a second Regular report for the
' same period, and Progressing after an advancement recommendation in the same grade.
' Local history may be incomplete or include alternative drafts, so both need review.
Private Sub RejectHistoryChecks(ByVal findings As Collection, ByVal f As Object, ByVal id As Long, ByVal kind As String, ByVal worn As String, ByVal recommendation As String)
    Dim member As Long, r As DAO.Recordset, starting As Date, ending As Date, regular As Boolean, progressing As Boolean
    Dim latest As Date, latestRecommendation As String, duplicate As String
    If id = 0 Or Not IsDate(RText(f, "PeriodStart")) Or Not IsDate(RText(f, "PeriodEnd")) Then Exit Sub
    regular = (RText(f, "ReportRegular") = "Yes")
    progressing = (recommendation = "Progressing" And (kind = "EVAL" Or kind = "CHIEFEVAL") And worn <> "")
    If Not regular And Not progressing Then Exit Sub
    member = Nz(MajorScalar("SELECT MemberID FROM MemberReport WHERE ReportID=" & id & " AND MatchState='Confirmed'"), 0)
    If member = 0 Then Exit Sub
    starting = DateValue(CDate(RText(f, "PeriodStart"))): ending = DateValue(CDate(RText(f, "PeriodEnd")))
    Set r = CurrentDb.OpenRecordset("SELECT r.ReportID, r.FormKind, r.ReportRegular, r.PeriodStart, r.PeriodEnd, r.RankRate, r.PromotionRecommendation FROM Report AS r INNER JOIN MemberReport AS m ON r.ReportID=m.ReportID WHERE m.MemberID=" & member & " AND m.MatchState='Confirmed' AND r.IsArchived=False AND r.ReportID<>" & id, dbOpenSnapshot)
    Do While Not r.EOF
        If Not IsNull(r!PeriodStart) And Not IsNull(r!PeriodEnd) Then
            If regular And Nz(r!ReportRegular, "") = "Yes" And Nz(r!FormKind, "") = kind Then
                If DateValue(r!PeriodStart) <= ending And DateValue(r!PeriodEnd) >= starting And Len(duplicate) = 0 Then duplicate = Format$(r!PeriodStart, "yyyy-mm-dd") & " to " & Format$(r!PeriodEnd, "yyyy-mm-dd")
            End If
            If progressing And DateValue(r!PeriodEnd) < starting And ValidationWornPaygrade(Nz(r!RankRate, "")) = worn Then
                If DateValue(r!PeriodEnd) > latest Then latest = DateValue(r!PeriodEnd): latestRecommendation = Nz(r!PromotionRecommendation, "")
            End If
        End If
        r.MoveNext
    Loop
    r.Close
    If Len(duplicate) > 0 Then VAdd findings, f, id, "PeriodStart", "Needs review", "RJ-DUPLICATE", "Another Regular report for this member covers part of the same period (" & duplicate & "). Only one Regular report is accepted for a period; archive the alternative draft or correct the dates."
    If latestRecommendation = "Promotable" Or latestRecommendation = "Must Promote" Or latestRecommendation = "Early Promote" Then VAdd findings, f, id, "PromotionRecommendation", "Needs review", "RJ-PROGRESSING", "This member's previous report in this grade recommended " & latestRecommendation & ". Progressing is not allowed while that recommendation is in effect; only Significant Problems withdraws it."
End Sub
