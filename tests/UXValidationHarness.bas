Attribute VB_Name = "UXValidationHarness"
Option Compare Binary
Option Explicit
Private count As Long
Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 6200, , message
    count = count + 1
End Sub
Private Function HasCode(ByVal issues As Collection, ByVal code As String) As Boolean
    Dim item As Object
    For Each item In issues
        If item("Code") = code Then HasCode = True: Exit Function
    Next
End Function
Public Function TestUXValidation() As String
    Dim id As Long, other As Long, groupID As Long, revision As Long, f As Object, result As Collection, item As Object, key As Variant, source As String, extra As Long, i As Long
    Dim changes As Object, runs As Long, operations As Long, revisions As Long, selected As Long, stage As String
    On Error GoTo Failed
    For Each key In Array("EVAL", "CHIEFEVAL", "FITREP")
        Set f = CreateObject("Scripting.Dictionary")
        f("FormKind") = CStr(key): f("PeriodStart") = "2026-01-01": f("PeriodEnd") = "2026-09-30"
        f("ObservationStatus") = "Observed": f("PromotionRecommendation") = "Promotable"
        f("UnitUIC") = "1234A": f("ReportingSeniorUIC") = "12345": f("CounselingDate") = "NOT REQ"
        Set result = ValidateLiveFields(f)
        Check Not HasCode(result, "UIC-FORMAT") And Not HasCode(result, "APP-DATE"), "Official UIC/counseling status rejected"
        Check HasCode(result, "CAREER-PRESENT") = (key <> "CHIEFEVAL"), "Form-specific career recommendation applicability"
        f("AssignmentRecommendationA") = "NONE": f("CounselingDate") = "NOT PERF"
        Set result = ValidateLiveFields(f): Check HasCode(result, "COUNSELOR"), "NOT PERF explanation missing"
        f("CounselorName") = "ILLNESS": Set result = ValidateLiveFields(f)
        Check Not HasCode(result, "COUNSELOR") And Not HasCode(result, "CAREER-PRESENT"), "Valid status explanation or NONE rejected"
        f("CounselingDate") = "2025-12-31": Set result = ValidateLiveFields(f)
        Check HasCode(result, "COUNSELING"), "Counseling outside period missed"
        f("CounselingDate") = "2026-05-01": Set result = ValidateLiveFields(f)
        Check Not HasCode(result, "COUNSELING"), "Counseling within period rejected"
        f("UnitUIC") = "AB345": Set result = ValidateLiveFields(f): Check HasCode(result, "UIC-FORMAT"), "Invalid first four UIC characters missed"
        f("UnitUIC") = "1234A": f("PromotionRecommendation") = "NOB"
        For i = 33 To 39: f("Trait" & i) = "NOB": Next
        For i = 33 To 36: f("Trait" & i) = "3.0": Next
        Set result = ValidateLiveFields(f): Check HasCode(result, "PARTIAL-NOB-LIMIT"), "Four partial NOB grades missed"
        f("Trait36") = "NOB": Set result = ValidateLiveFields(f)
        Check Not HasCode(result, "PARTIAL-NOB-LIMIT") And Not HasCode(result, "PARTIAL-NOB-MARKS"), "Three partial NOB grades rejected"
        f("Trait33") = "1.0": Set result = ValidateLiveFields(f)
        Check HasCode(result, "PARTIAL-NOB-ADVERSE") And HasCode(result, "JUSTIFY-TRAITS") And HasCode(result, "ADVERSE-REFERRAL"), "Adverse partial NOB missed"
        f("Trait33") = "3.0": f("Trait39") = "": Set result = ValidateLiveFields(f)
        Check HasCode(result, "PARTIAL-NOB-MARKS"), "Blank partial NOB trait missed"
        f("PromotionRecommendation") = "Significant Problems": Set result = ValidateLiveFields(f)
        Check HasCode(result, "JUSTIFY-ADVERSE"), "Adverse recommendation justification missed"
        f("ReportRegular") = "Yes": f("PeriodStart") = "2024-01-01": Set result = ValidateLiveFields(f)
        Check HasCode(result, "REGULAR-LENGTH"), "Extended Regular approval check missed"
        f("ReportRegular") = "No": Set result = ValidateLiveFields(f)
        Check Not HasCode(result, "REGULAR-LENGTH"), "Regular duration check applied to other type"
    Next
    ConfigureWorkspace "SYNTHETIC VALIDATION TEST", "Validation tester"
    id = CommitDraft(0, "EVAL", "Validation test", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test")
    Set f = SavedFields(id): Set result = ValidateLiveFields(f, id)
    Check HasCode(result, "SSN-PRESENT"), "Missing SSN not listed"
    Check HasCode(result, "NAME-PRESENT"), "Missing official name not listed"
    For Each key In Array("NAME-PRESENT", "RATE-PRESENT", "UIC-PRESENT", "SSN-PRESENT", "COMPONENT", "REPORT-TYPE", "WHOLE-NOB", "TRAIT-RECOMMENDATION", "TRAIT-TWO", "JUNIOR-OFFICER", "GROUP-OFFICER", "GROUP-ENLISTED", "GROUP-CATEGORY", "CAPACITY")
        source = ValidationPolicyReference(CStr(key))
        Check InStr(source, "BUPERSINST 1610.10H CH-2") > 0 And InStr(source, "printed page") > 0 And InStr(source, "paragraph 1-2") > 0 And InStr(source, "Excerpt:") > 0 And InStr(source, NAVY_RULE_SOURCE_HASH) > 0, "Source detail incomplete: " & key
    Next
    Check InStr(ValidationPolicyReference("APP-DIGITS"), "Application rule") > 0, "Storage rule mislabeled as policy"
    Check InStr(ValidationPolicyReference("UNKNOWN-POLICY"), "not been verified") > 0, "Unverified policy received an invented citation"
    For Each key In Array(Array("IT2", "E5"), Array("ET2", "E5"), Array("ITCS", "E8"), Array("ITCM", "E9"), Array("ITC", "E7"), Array("AN", "E3"), Array("SA", "E2"), Array("SR", "E1"), Array("LCDR", "O4"), Array("CWO3", "W3"))
        Check ValidationWornPaygrade(CStr(key(0))) = CStr(key(1)), "Worn-grade mapping: " & key(0)
    Next
    Check ValidationWornPaygrade("Unknown") = "", "Unrecognized rate should need review"
    f("ComponentActive") = "Yes": f("ComponentTAR") = "No": f("ComponentInactive") = "No": f("ComponentATADSW") = "No"
    f("ReportRegular") = "Yes": f("ReportConcurrent") = "Yes": f("ReportOperationalCommander") = "No"
    f("ObservationStatus") = "Observed": f("PromotionRecommendation") = "Early Promote": f("Trait33") = "2.0"
    Set result = ValidateLiveFields(f, id)
    Check Not HasCode(result, "REPORT-TYPE"), "Concurrent/Regular combination wrongly rejected"
    Check HasCode(result, "TRAIT-TWO"), "2.0 / EP conflict missing"
    f("OccasionPeriodic") = "Yes": f("OccasionSpecial") = "Yes": Set result = ValidateLiveFields(f, id)
    Check HasCode(result, "OCCASION"), "Special combined with Periodic not found"
    f("OccasionSpecial") = "No"
    f("Trait33") = "1.0": Set result = ValidateLiveFields(f, id)
    Check HasCode(result, "TRAIT-RECOMMENDATION"), "1.0 / EP conflict missing"
    f("Trait33") = "3.0": f("PromotionRecommendation") = "NOB": Set result = ValidateLiveFields(f, id)
    Check HasCode(result, "PARTIAL-NOB") And Not HasCode(result, "WHOLE-NOB"), "Partial NOB treated as wholly NOB"
    f("ObservationStatus") = "NOB": Set result = ValidateLiveFields(f, id)
    Check HasCode(result, "WHOLE-NOB"), "Whole NOB graded trait not found"
    f("FormKind") = "FITREP": f("ObservationStatus") = "Observed": f("RankRate") = "ENS": f("Designator") = "1110": f("PromotionRecommendation") = "Early Promote"
    Set result = ValidateLiveFields(f, id): Check HasCode(result, "JUNIOR-OFFICER"), "ENS restriction missing"
    f("Designator") = "6110": Set result = ValidateLiveFields(f, id): Check Not HasCode(result, "JUNIOR-OFFICER"), "LDO exception missing"
    f("FormKind") = "CHIEFEVAL": f("PromotionRecommendation") = "Promotable": f("Trait37") = "2.0": Set result = ValidateLiveFields(f, id)
    Check HasCode(result, "TRAIT-CRITICAL"), "Chief accountability restriction missing"
    f("Trait33") = "2.0": f("Trait34") = "2.0": Set result = ValidateLiveFields(f, id)
    Check HasCode(result, "TRAIT-TWO-COUNT"), "Three 2.0 traits / Promotable not found"
    f("MemberSSN") = "123": Set result = ValidateLiveFields(f, id): Check HasCode(result, "APP-DIGITS"), "Invalid SSN missing"
    stage = "native report"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    revision = CurrentRevision(id): runs = DCount("*", "ValidationRun"): revisions = DCount("*", "ReportRevision"): operations = DCount("*", "OperationJournal")
    Forms!frmEditor!txtMemberFirstName = "Pending first name"
    Check EditorPendingFields()("MemberFirstName") = "Pending first name", "Pending snapshot lost field"
    Check EditorPendingFields().count = UBound(DraftFieldKeys()) + 1, "Pending snapshot lost fields"
    Forms!frmEditor!uxFieldSurface.SetFocus
    Forms!frmEditor!uxFieldSurface.Form!txtMemberSSN.SetFocus
    Forms!frmEditor!uxFieldSurface.Form!txtMemberSSN.Text = "321"
    Check EditorPendingFields()("MemberSSN") = "321", "Active text before AfterUpdate omitted"
    Check UXLiveTyped(), "Live typed paint failed"
    Check InStr(Forms!frmEditor!uxDirty.Caption, "issues") > 0, "Live count missing"
    UXOpenValidation
    Check Forms!frmUXValidation!lstIssues.ListCount > 0, "Report issue list empty"
    For i = 0 To Forms!frmUXValidation!lstIssues.ListCount - 1
        Forms!frmUXValidation!lstIssues = Forms!frmUXValidation!lstIssues.ItemData(i): UXValidationSelection
        If InStr(Forms!frmUXValidation!txtEvidence, "Excerpt:") > 0 Then Exit For
    Next
    Check InStr(Forms!frmUXValidation!txtEvidence, "Excerpt:") > 0, "Native policy issue source missing"
    selected = Forms!frmUXValidation!lstIssues
    UXLoadValidation
    Check Forms!frmUXValidation!lstIssues = selected, "Refresh reset selection"
    Check UXValidationOpenField(), "Open affected field failed"
    Check Forms!frmEditor!txtMemberFirstName = "Pending first name", "Issue navigation lost pending values"
    Check CurrentRevision(id) = revision And DCount("*", "ReportRevision") = revisions And DCount("*", "ValidationRun") = runs And DCount("*", "OperationJournal") = operations, "Live checks wrote database state"
    UXValidationClose
    stage = "group"
    other = CommitDraft(0, "EVAL", "Other test", "Other member", #1/1/2026#, #8/31/2026#, "", "Test")
    groupID = CreateSummaryGroup("Synthetic validation group", "EVAL", "E5")
    AddGroupMember groupID, id, "E5", "": AddGroupMember groupID, other, "E5", ""
    Set result = ValidateLiveGroup(groupID)
    Check HasCode(result, "CAREER-PRESENT") And HasCode(result, "COUNSELING"), "Expanded report checks omitted from summary group"
    Check HasCode(result, "GROUP-ENLISTED"), "Group eligibility differences missing"
    source = ""
    For Each item In result
        If item("ReportID") = id And item("Key") = "MemberFirstName" Then source = "Unexpected"
    Next
    Check source = "", "Optional name parts treated as required"
    Forms!frmEditor!txtMemberSSN = "123"
    Set result = ValidateLiveGroup(groupID): Check HasCode(result, "APP-DIGITS"), "Group omitted pending member value"
    DoCmd.OpenForm "frmUXValidation", , , , , , "G:" & groupID
    Check Forms!frmUXValidation!lstIssues.ListCount > 0 And InStr(Forms!frmUXValidation!lblBasis.Caption, "All group members") > 0, "Group issue view failed"
    'On initial open this automation host can report a zero client width.
    'Use the native documented sizing properties; do not treat it as a render.
    Set f = Forms!frmUXValidation
    For Each key In Array(6500, 14000, 17000, 6500)
        f.InsideWidth = CLng(key): f.InsideHeight = 9500
        Check UXValidationLayout(), "Native validation layout failed after native sizing"
        Check f!txtEvidence.Top + f!txtEvidence.Height + 250 <= f!cmdOpen.Top, "Evidence overlaps validation actions"
        Check f!lstIssues.Top + f!lstIssues.Height < f!lblEvidence.Top, "Findings overlap evidence label"
        Check f!cmdWindowHelp.Left >= f!lblHeading.Left + f!lblHeading.Width, "Help overlaps validation heading"
        Check f!cmdBack.Left + f!cmdBack.Width <= f.Width, "Close validation action outside form"
        Check f!lstIssues.Controls("lblIssues").Caption = "&Findings", "Findings has no attached accessible label"
        Check f!txtEvidence.Controls("lblEvidence").Caption = "&Evidence and source", "Evidence has no attached accessible label"
        Check f!lstIssues.TabIndex < f!txtEvidence.TabIndex And f!txtEvidence.TabIndex < f!cmdOpen.TabIndex, "Validation tab sequence is not findings/evidence/action"
        f!lstIssues.SetFocus: Check Screen.ActiveControl.Name = "lstIssues", "Cannot focus findings"
        f!txtEvidence.SetFocus: Check Screen.ActiveControl.Name = "txtEvidence", "Cannot focus read-only evidence"
    Next
    UXValidationClose
    For i = 1 To 3
        extra = CommitDraft(0, "EVAL", "Pagination member " & i, "Synthetic page " & i, #1/1/2026#, #9/30/2026#, "", "Test")
        AddGroupMember groupID, extra, "E5", ""
    Next
    DoCmd.OpenForm "frmUXValidation", , , , , , "G:" & groupID
    Check Forms!frmUXValidation!cmdNext.Enabled, "More than 50 findings lost"
    Check UXValidationPage(1) And Forms!frmUXValidation!lstIssues > 50, "Next findings page failed"
    Check UXValidationPage(-1) And Forms!frmUXValidation!lstIssues <= 50, "Previous findings page failed"
    stage = "close validation": UXValidationClose
    stage = "close editor": UXLeaveEditor "Discard"
    TestUXValidation = "PASS: " & count & " live validation and source assertions": Exit Function
Failed:
    TestUXValidation = "FAIL after " & count & " (" & stage & "): " & Err.Number & " / " & Err.Description
End Function
