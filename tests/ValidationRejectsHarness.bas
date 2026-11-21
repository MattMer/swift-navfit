Attribute VB_Name = "ValidationRejectsHarness"
Option Explicit
' Checks drawn from PERS-32 reject reasons (modValidationRejects): each rule found and
' not found, the stated exceptions, member history, and every cited excerpt found word
' for word in the built-in EVALMAN.
Private n As Long

Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 8600, , message
    n = n + 1
End Sub

Private Function Finding(ByVal findings As Collection, ByVal code As String, Optional ByVal field As String = "") As Object
    Dim item As Object
    For Each item In findings
        If item("Code") = code And (field = "" Or item("Key") = field) Then Set Finding = item: Exit Function
    Next
End Function

Private Function Has(ByVal findings As Collection, ByVal code As String, Optional ByVal field As String = "") As Boolean
    Has = Not Finding(findings, code, field) Is Nothing
End Function

Private Function RejectCount(ByVal findings As Collection) As String
    Dim item As Object
    For Each item In findings
        If Left$(item("Code"), 3) = "RJ-" Then RejectCount = RejectCount & item("Code") & " "
    Next
End Function

Private Function Spaces(ByVal text As String) As String
    text = Replace(Replace(Replace(text, vbCr, " "), vbLf, " "), vbTab, " ")
    Do While InStr(text, "  ") > 0: text = Replace(text, "  ", " "): Loop
    Spaces = Trim$(text)
End Function

Private Function TraitKeyTitled(ByVal kind As String, ByVal word As String) As String
    Dim key As Variant
    For Each key In TraitKeys()
        If InStr(UCase$(CatalogueTraitTitle(kind, CLng(Mid$(CStr(key), 6)))), word) > 0 Then TraitKeyTitled = CStr(key): Exit Function
    Next
End Function

' A complete report with nothing for these checks to find.
Private Function Example(ByVal id As Long, ByVal kind As String) As Object
    Dim f As Object, key As Variant
    Set f = SavedFields(id)
    f("FormKind") = kind: f("OfficialName") = "EXAMPLE, A": f("RankRate") = IIf(kind = "EVAL", "IT1", IIf(kind = "CHIEFEVAL", "ITC", "LT"))
    f("MemberSSN") = "000000001": f("ReportingSeniorSSN") = "000000002": f("ReportingSeniorName") = "SENIOR, A"
    f("ReportingSeniorGrade") = "CDR": f("ReportingSeniorDesignator") = "1310": f("ReportingSeniorTitle") = "CO"
    f("ReportingSeniorUIC") = "12345": f("UnitUIC") = "12345": f("ShipStation") = "SYNTHETIC COMMAND"
    f("PeriodStart") = #1/1/2026#: f("PeriodEnd") = #9/30/2026#: f("DateReported") = "2026-01-01": f("PromotionStatus") = "REGULAR"
    f("ComponentActive") = "Yes": f("ReportRegular") = "Yes": f("OccasionDetachment") = "Yes": f("OccasionPeriodic") = "No"
    f("OccasionFrocking") = "No": f("OccasionSpecial") = "No": f("ReportConcurrent") = "No": f("ReportOperationalCommander") = "No"
    f("ObservationStatus") = "Observed": f("PromotionRecommendation") = "Promotable": f("Designator") = IIf(kind = "FITREP", "1110", "AW")
    For Each key In TraitKeys(): f(CStr(key)) = "4.0": Next
    f("RetentionYes") = "Yes": f("RetentionNo") = "No": f("OfficialComments") = "Synthetic comments."
    f("ReadinessSummary") = "P": f("ReadinessDate") = "2026-06-01": f("BilletSubcategory") = "NA"
    f("PrimaryDuty") = "TECHNICIAN": f("DutyDescription") = "Synthetic duty description.": f("CommandEmployment") = "Synthetic command employment."
    Set Example = f
End Function

Private Sub CheckReferences()
    Dim code As Variant, section As Long, excerpt As String, reference As String
    For Each code In Array("RJ-SUBCAT-BLANK", "RJ-SUBCAT-DUTY", "RJ-FROCK-STATUS", "RJ-FROCK-OFFICER", "RJ-ENS-SELECTED", "RJ-SPECIAL", "RJ-OPS-CDR-MEMBER", "RJ-OPS-CDR-TITLE", "RJ-ADVERSE-SIGN", "RJ-PFA-FAILURES", "RJ-PROGRESSING", "RJ-DATE-REPORTED", "RJ-FUTURE-END", "RJ-DUPLICATE")
        section = EvalmanSectionByTitle(RejectRuleSection(CStr(code)))
        Check section >= 0, code & " names no EVALMAN section"
        Check EvalmanSectionForFinding(CStr(code), "OfficialName") = section, code & " opens a different EVALMAN section from the one it cites"
        excerpt = RejectRuleExcerpt(CStr(code)): reference = RejectPolicyReference(CStr(code))
        Check InStr(reference, "PERS-32 reject reason ") = 1, code & " reference does not name its PERS-32 reason"
        If Len(excerpt) > 0 Then
            Check InStr(1, Spaces(EvalmanText(section)), Spaces(excerpt), vbBinaryCompare) > 0, code & " excerpt is not in EVALMAN " & EvalmanHeader(section, 1)
            Check InStr(reference, EvalmanSource()) > 0 And InStr(reference, excerpt) > 0, code & " reference lacks its source or excerpt"
        Else
            Check InStr(reference, "EVALMAN has no sentence") > 0, code & " web-only reason claims an EVALMAN excerpt"
        End If
    Next
End Sub

Public Function TestValidationRejects() As String
    Dim stage As String, id As Long, kind As Variant, f As Object, result As Collection, key As String, item As Object
    Dim older As Long, overlap As Long, member As Long, changes As Object, trait As Variant
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC REJECT CHECKS", "Tester"
    id = CommitDraft(0, "EVAL", "Synthetic", "Example", Null, Null, "")
    stage = "references": CheckReferences

    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        stage = "baseline " & kind
        Set result = ValidateLiveFields(Example(id, CStr(kind)), 0)
        Check RejectCount(result) = "", kind & " complete report has reject findings: " & RejectCount(result)

        stage = "block 21 " & kind
        Set f = Example(id, CStr(kind)): f("BilletSubcategory") = Null
        Set item = Finding(ValidateLiveFields(f, 0), "RJ-SUBCAT-BLANK", "BilletSubcategory")
        Check Not item Is Nothing, kind & " blank block 21 not found"
        Check item("Severity") = "Incomplete draft", kind & " blank block 21 severity " & item("Severity")
        f("BilletSubcategory") = "STUDENT": Check Has(ValidateLiveFields(f, 0), "RJ-SUBCAT-DUTY"), kind & " STUDENT without block 29 entry not found"
        f("PrimaryDuty") = "STUDENT": Check Not Has(ValidateLiveFields(f, 0), "RJ-SUBCAT-DUTY"), kind & " STUDENT matching block 29 flagged"
        f("BilletSubcategory") = "CO AFLOAT": f("PrimaryDuty") = "COMMAND STAFF": Check Has(ValidateLiveFields(f, 0), "RJ-SUBCAT-DUTY"), kind & " CO AFLOAT without CO duty not found"
        f("PrimaryDuty") = "CO": Check Not Has(ValidateLiveFields(f, 0), "RJ-SUBCAT-DUTY"), kind & " CO AFLOAT with CO duty flagged"

        stage = "adverse " & kind
        Set f = Example(id, CStr(kind)): f("Trait33") = "1.0": f("PromotionRecommendation") = "Significant Problems"
        Check Has(ValidateLiveFields(f, 0), "RJ-ADVERSE-SIGN"), kind & " adverse report signature not flagged"

        stage = "special " & kind
        Set f = Example(id, CStr(kind)): f("OccasionDetachment") = "No": f("OccasionSpecial") = "Yes": f("OfficialComments") = Null
        Set item = Finding(ValidateLiveFields(f, 0), "RJ-SPECIAL")
        Check Not item Is Nothing, kind & " Special without comments not found"
        Check item("Severity") = "Incomplete draft", kind & " Special without comments severity " & item("Severity")
        f("OfficialComments") = "Submitted for the synthetic promotion board."
        Set item = Finding(ValidateLiveFields(f, 0), "RJ-SPECIAL")
        Check item("Severity") = "Needs review", kind & " Special with comments severity " & item("Severity")

        stage = "dates " & kind
        Set f = Example(id, CStr(kind)): f("DateReported") = "2026-10-01"
        Check Has(ValidateLiveFields(f, 0), "RJ-DATE-REPORTED", "DateReported"), kind & " date reported after the period not found"
        f("DateReported") = "2026-09-30": Check Not Has(ValidateLiveFields(f, 0), "RJ-DATE-REPORTED"), kind & " date reported on the end date flagged"
        f("PeriodEnd") = DateAdd("d", 30, Date): f("DateReported") = "2026-01-01"
        Set item = Finding(ValidateLiveFields(f, 0), "RJ-FUTURE-END")
        Check Not item Is Nothing, kind & " future end date not found"
        Check item("Severity") = "Information", kind & " future end date severity " & item("Severity")

        stage = "wholly NOB " & kind
        Set f = Example(id, CStr(kind)): f("ObservationStatus") = "NOB": f("PromotionRecommendation") = "NOB": f("ReadinessSummary") = "F"
        For Each trait In TraitKeys(): f(CStr(trait)) = "NOB": Next
        Set result = ValidateLiveFields(f, 0)
        Check Not Has(result, "RJ-PFA-FAILURES") And Not Has(result, "RJ-ADVERSE-SIGN"), kind & " wholly NOB report flagged: " & RejectCount(result)
    Next

    stage = "E-7 to E-9 block 37"
    'EVALMAN's E-7 to E-9 Character trait is block 37, Accountability, on the current form.
    key = TraitKeyTitled("CHIEFEVAL", "ACCOUNTABILITY"): Check key = "Trait37", "CHIEFEVAL Accountability is " & key
    Set f = Example(id, "CHIEFEVAL"): f(key) = "2.0"
    Set result = ValidateLiveFields(f, 0)
    Check Has(result, "TRAIT-CRITICAL", key), "CHIEFEVAL Accountability 2.0 with Promotable not found"
    Check Has(result, "RJ-ADVERSE-SIGN") And Has(result, "ADVERSE-REFERRAL"), "CHIEFEVAL Accountability 2.0 not treated as adverse"
    Set f = Example(id, "CHIEFEVAL"): f("Trait33") = "2.0"
    Check Not Has(ValidateLiveFields(f, 0), "RJ-ADVERSE-SIGN"), "One 2.0 outside block 37 treated as adverse"

    stage = "frocking"
    Set f = Example(id, "EVAL"): f("OccasionDetachment") = "No": f("OccasionFrocking") = "Yes"
    Set result = ValidateLiveFields(f, 0)
    Check Has(result, "RJ-FROCK-STATUS", "PromotionStatus"), "Frocking report with REGULAR in block 8 not found"
    Check Not Has(result, "RJ-FROCK-OFFICER"), "Frocking to E-7 questioned"
    f("PromotionStatus") = "SELECTED": Check Not Has(ValidateLiveFields(f, 0), "RJ-FROCK-STATUS"), "Frocking report with SELECTED flagged"
    Set f = Example(id, "FITREP"): f("OccasionDetachment") = "No": f("OccasionFrocking") = "Yes": f("PromotionStatus") = "SELECTED"
    Check Has(ValidateLiveFields(f, 0), "RJ-FROCK-OFFICER"), "Officer frocking report not questioned"

    stage = "ensign"
    Set f = Example(id, "FITREP"): f("RankRate") = "ENS": f("PromotionStatus") = "SELECTED"
    Check Has(ValidateLiveFields(f, 0), "RJ-ENS-SELECTED", "PromotionStatus"), "Ensign with SELECTED not found"
    'Promoting to LTJG: REGULAR, Special, PROMOTING TO LTJG first in the comments.
    f("PromotionStatus") = "REGULAR": f("OccasionDetachment") = "No": f("OccasionSpecial") = "Yes": f("OfficialComments") = "PROMOTING TO LTJG. Synthetic comments."
    Set result = ValidateLiveFields(f, 0)
    Check Not Has(result, "RJ-ENS-SELECTED") And Not Has(result, "RJ-FROCK-STATUS") And Not Has(result, "RJ-FROCK-OFFICER"), "Ensign promoting to LTJG flagged: " & RejectCount(result)
    Check Finding(result, "RJ-SPECIAL")("Severity") = "Needs review", "Ensign promoting to LTJG Special not left to review"

    stage = "operational commander"
    Set f = Example(id, "FITREP"): f("ReportRegular") = "No": f("ReportOperationalCommander") = "Yes": f("ReportingSeniorTitle") = "XO"
    Set result = ValidateLiveFields(f, 0)
    Check Has(result, "RJ-OPS-CDR-MEMBER", "PrimaryDuty") And Has(result, "RJ-OPS-CDR-TITLE", "ReportingSeniorTitle"), "Ops Cdr report on a non-CO by a non-commander not found: " & RejectCount(result)
    f("PrimaryDuty") = "CO": f("ReportingSeniorTitle") = "COMDESRON"
    Set result = ValidateLiveFields(f, 0)
    Check Not Has(result, "RJ-OPS-CDR-MEMBER") And Not Has(result, "RJ-OPS-CDR-TITLE"), "Ops Cdr report on a CO by a commander flagged"

    stage = "physical readiness"
    For Each kind In Array("EVAL", "FITREP")
        Set f = Example(id, CStr(kind)): f("ReadinessSummary") = "PF"
        Check Has(ValidateLiveFields(f, 0), "RJ-PFA-FAILURES", "ReadinessSummary"), kind & " PFA failure not flagged"
        f(TraitKeyTitled(CStr(kind), "BEARING")) = "1.0"
        Check Not Has(ValidateLiveFields(f, 0), "RJ-PFA-FAILURES"), kind & " PFA failure with Military Bearing 1.0 flagged"
    Next

    stage = "member history"
    older = CommitDraft(0, "EVAL", "Synthetic older", "HISTORY, MEMBER", #1/1/2025#, #12/31/2025#, "")
    overlap = CommitDraft(0, "EVAL", "Synthetic overlap", "HISTORY, MEMBER", #3/1/2026#, #6/30/2026#, "")
    id = CommitDraft(0, "EVAL", "Synthetic current", "HISTORY, MEMBER", #1/1/2026#, #9/30/2026#, "")
    Set changes = CreateObject("Scripting.Dictionary")
    changes("RankRate") = "IT1": changes("ReportRegular") = "Yes": changes("PromotionRecommendation") = "Promotable"
    Call CommitSavedFields(older, CurrentRevision(older), changes, "Synthetic history")
    changes("PromotionRecommendation") = "Must Promote"
    Call CommitSavedFields(overlap, CurrentRevision(overlap), changes, "Synthetic history")
    member = Nz(DLookup("MemberID", "MemberReport", "ReportID=" & older), 0): Check member > 0, "Synthetic report has no directory member"
    CurrentDb.Execute "UPDATE MemberReport SET MemberID=" & member & ", MatchState='Confirmed' WHERE ReportID In (" & older & "," & overlap & "," & id & ")", dbFailOnError
    Set f = Example(id, "EVAL"): f("PromotionRecommendation") = "Progressing"
    Set result = ValidateLiveFields(f, id)
    Check Has(result, "RJ-DUPLICATE", "PeriodStart"), "Overlapping Regular report not found"
    Check Has(result, "RJ-PROGRESSING", "PromotionRecommendation"), "Progressing after Promotable in the same grade not found"
    ArchiveDraft overlap, True
    f("PromotionRecommendation") = "Promotable": Set result = ValidateLiveFields(f, id)
    Check Not Has(result, "RJ-DUPLICATE") And Not Has(result, "RJ-PROGRESSING"), "Archived or answered history still flagged: " & RejectCount(result)
    changes.RemoveAll: changes("PromotionRecommendation") = "Significant Problems"
    Call CommitSavedFields(older, CurrentRevision(older), changes, "Synthetic withdrawal")
    f("PromotionRecommendation") = "Progressing": Check Not Has(ValidateLiveFields(f, id), "RJ-PROGRESSING"), "Progressing after Significant Problems flagged"

    TestValidationRejects = "PASS: " & n & " reject-reason checks, exceptions, member history and EVALMAN excerpts"
    Exit Function
Failed:
    TestValidationRejects = "FAIL: " & stage & ": " & Err.Number & " " & Err.Description
End Function
