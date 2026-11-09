Attribute VB_Name = "RestrictedContentHarness"
Option Explicit
'Restricted-content check: every rule flags its synthetic examples, ordinary Navy
'wording (and the NavyWriter catalogue) stays clear, dismissals and allowed phrases
'hide findings and come back when undone, and a full pass stays fast.
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 7895, , message
    checks = checks + 1
End Sub
'Rule IDs found in one block, sorted and joined with |.
Private Function Flags(ByVal text As String, Optional ByVal key As String = "OfficialComments", Optional ByVal kind As String = "EVAL", Optional ByVal reportID As Long = 0) As String
    Dim item As Variant, seen As Object, k As Variant, list() As String, i As Long, j As Long, t As String
    Set seen = CreateObject("Scripting.Dictionary")
    For Each item In RestrictedMatches(kind, key, text, reportID): seen(CStr(item(0))) = True: Next
    If seen.Count = 0 Then Exit Function
    ReDim list(0 To seen.Count - 1)
    For Each k In seen.Keys: list(i) = k: i = i + 1: Next
    For i = 0 To UBound(list) - 1
        For j = i + 1 To UBound(list)
            If list(j) < list(i) Then t = list(i): list(i) = list(j): list(j) = t
        Next
    Next
    Flags = Join(list, "|")
End Function
Private Sub Flagged(ByVal rule As String, ByVal text As String, Optional ByVal key As String = "OfficialComments")
    Check InStr("|" & Flags(text, key) & "|", "|" & rule & "|") > 0, rule & " missed: " & text & " -> " & Flags(text, key)
End Sub
Private Sub Clear(ByVal text As String, Optional ByVal key As String = "OfficialComments")
    Check Flags(text, key) = "", "False positive " & Flags(text, key) & ": " & text
End Sub
Private Function FindingCount(ByVal id As Long, ByVal code As String) As Long
    Dim item As Variant
    For Each item In ValidateLiveFields(SavedFields(id), id)
        If item("Code") = code Then FindingCount = FindingCount + 1
    Next
End Function
Public Function TestRestrictedContent() As String
    Dim text As String, i As Long, id As Long, other As Long, changes As Object, dismissal As Long, exception As Long, started As Double, elapsed As Double, sweep As String, item As Variant, found As Boolean
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC RESTRICTED CONTENT", "Synthetic Reviewer"
    Check RestrictedRuleCount() >= 19 And Len(RestrictedCatalogueHash()) = 64, "Rule catalogue missing"

    ' Each rule flags its synthetic examples.
    Flagged "RC-FOS", "Despite his previous failure of selection, he excelled as division officer."
    Flagged "RC-FOS", "Was passed over for LCDR but remained focused on the mission."
    Flagged "RC-FOS", "Not selected by the FY26 CPO board, he redoubled his efforts."
    Flagged "RC-RESIGNATION", "Submitted his resignation in March and trained his relief."
    Flagged "RC-RESIGNATION", "Resigned his commission effective June."
    Flagged "RC-PROCEEDINGS", "Pending NJP for an alleged DUI, he continued to perform."
    Flagged "RC-PROCEEDINGS", "Currently under investigation for misconduct."
    Flagged "RC-PROCEEDINGS", "Awaiting court-martial, he still stood every watch."
    Flagged "RC-EXONERATED", "Acquitted at court-martial in June, he returned to full duty."
    Flagged "RC-EXONERATED", "The allegation against her was unsubstantiated."
    Flagged "RC-CENSURE", "Received a non-punitive letter of caution for tardiness."
    Flagged "RC-CENSURE", "Issued an NPLOC for repeated lateness."
    Flagged "RC-CENSURE", "Received a Page 13 for a liberty violation."
    Flagged "RC-INVESTIGATION", "He was the subject of a command investigation."
    Flagged "RC-INVESTIGATION", "The investigation found he followed every procedure."
    Flagged "RC-FAMILY", "Despite his wife's illness, he never missed a watch."
    Flagged "RC-FAMILY", "Balanced demanding duties as a single parent."
    Flagged "RC-FAMILY", "Supported his family through a difficult move."
    Flagged "RC-FAMILY", "Recently married, she still led the division."
    Flagged "RC-MEDICAL", "Returned from maternity leave and immediately led the division."
    Flagged "RC-MEDICAL", "While on LIMDU, he trained 20 Sailors."
    Flagged "RC-MEDICAL", "Overcame PTSD to qualify as EOOW."
    Flagged "RC-MEDICAL", "Despite her pregnancy, she qualified in record time."
    Flagged "RC-MEDICAL", "Recovering from his surgery, he managed the tool room."
    Flagged "RC-MEDICAL", "Managed his diabetes while deployed."
    Flagged "RC-COURTS", "Served as a court-martial member and was highly effective."
    Flagged "RC-COURTS", "Sat on an administrative separation board."
    Flagged "RC-AWARD-RECOMMENDED", "Recommended for a Navy and Marine Corps Achievement Medal."
    Flagged "RC-AWARD-RECOMMENDED", "His NAM is pending approval."
    Flagged "RC-AWARD-RECOMMENDED", "Submitted for Sailor of the Year."
    Flagged "RC-CIVILIAN", "He will excel in any civilian career."
    Flagged "RC-CIVILIAN", "His employer praised his reserve service."
    Flagged "RC-CIVILIAN", "A future leader in the private sector."
    Flagged "RC-DFC", "Detached for cause in May."
    Flagged "RC-DFC", "Relieved for cause after the inspection."
    Flagged "RC-CLASSIFIED", "Supported (S) operations in the region."
    Flagged "RC-CLASSIFIED", "Briefed SECRET//NOFORN material to the staff."
    Flagged "RC-REPRISAL", "Filed an EO complaint against his chain of command."
    Flagged "RC-REPRISAL", "A victim of sexual assault, she persevered."
    Flagged "RC-REPRISAL", "Made a protected communication to his congressman."
    Flagged "RC-PROTECTED", "His Muslim faith guides his leadership."
    Flagged "RC-PROTECTED", "At 52, his age never slowed him down."
    Flagged "RC-POLITICAL", "Campaigned for a Republican candidate in his town."
    Flagged "RC-PII", "SSN 123-45-6789 is on file."
    Flagged "RC-PII", "Reach him at 555-123-4567 or john.doe@example.com."
    Flagged "RC-CAPS", "LED THE DIVISION THROUGH A DEMANDING INSPECTION CYCLE WITH ZERO DISCREPANCIES."
    Flagged "RC-ANTICIPATED", "Will complete his associate degree in May.", "Qualifications"
    Flagged "RC-ANTICIPATED", "Currently enrolled in the CPO 365 course.", "Qualifications"
    Check InStr(Flags("Will complete his associate degree in May."), "RC-ANTICIPATED") = 0, "Block 44 rule ran outside block 44"
    Check Flags("Despite his wife's illness, he stood watch.", "OfficialComments", "FITREP") <> "" And Flags("Despite his wife's illness, he stood watch.", "Trait33Comment1", "CHIEFEVAL") <> "", "Rules not applied on every report type and block"

    ' Ordinary Navy wording stays clear.
    Clear "Led 12 Sailors through a demanding inspection, earning zero discrepancies."
    Clear "Served as Family Readiness Group liaison, supporting 300 families."
    Clear "Screened 400 Sailors for hypertension and treated 50 patients at sick call."
    Clear "Supported builder's trials and sea trials pending certification."
    Clear "Processed 45 award recommendations for Sailors in the department."
    Clear "Received the Navy and Marine Corps Achievement Medal for superior performance."
    Clear "Diagnosed and repaired faults in the SPS-67 radar system."
    Clear "Coordinated SAPR training as the command Victim Advocate."
    Clear "Sailor of the Quarter. Strongly recommended for advancement; my highest recommendation."
    Clear "Managed the command suicide prevention program with zero gaps."
    Clear "Mentored 30 children at a local elementary school."
    Clear "Organized the command family day picnic for 400 guests."
    Clear "Qualified search and rescue (SAR) swimmer."
    Clear "Completed 120 hours of training in 2025 and 4,500 maintenance actions."
    Clear "Voting assistance officer; registered 140 Sailors."
    Clear "Executed the liberal leave policy during the holiday stand-down."
    Clear "Supervised 20 Sailors through the IG inspection with no findings."
    Clear "Drafted the exercise letter of instruction for COMPTUEX."
    Clear "Expected to deploy in the spring, he trained 30 Sailors.", "OfficialComments"
    Clear "Earned his associate degree and the ESWS pin.", "Qualifications"
    Clear "Led the effort to reduce injuries, achieving zero lost workdays."
    Clear "Maintained medical records for 600 Sailors with 100 percent accuracy."
    Clear "Senior enlisted leader for the electronics division (ET, IT and FC ratings)."
    Clear "Applied lessons from the Battle of Midway in a professional development race against time."

    ' The NavyWriter catalogue is ordinary writing: none of it should be flagged.
    For i = 0 To NavyWriterCount() - 1
        text = NavyWriterValue(i, 1)
        If Flags(text) <> "" Then sweep = sweep & vbCrLf & Flags(text) & ": " & text
    Next
    Check sweep = "", "NavyWriter entries flagged:" & sweep

    ' Validation shows the finding with its citation; Dismiss hides it for this report only.
    id = CommitDraft(0, "EVAL", "Content check", "CONTENT, CHECK", Null, Null, "", "Content check", "Despite his wife's illness, he never missed a watch.")
    other = CommitDraft(0, "EVAL", "Content check two", "CONTENT, OTHER", Null, Null, "", "Content check", "His wife hosted the command ombudsman meeting.")
    Check FindingCount(id, "RC-FAMILY") = 1 And FindingCount(other, "RC-FAMILY") = 1, "Family finding not reported"
    For Each item In ValidateLiveFields(SavedFields(id), id)
        If item("Code") = "RC-FAMILY" Then
            found = True
            Check item("Severity") = "Needs review" And item("Match") = "wife's" And item("Key") = "OfficialComments", "Finding details wrong: " & item("Message")
            Check InStr(item("Reference"), "13-5.f") > 0 And InStr(item("Reference"), "PDF page 129") > 0, "Citation missing: " & item("Reference")
            Check Mid$(SavedFields(id)("OfficialComments"), item("Start"), item("Length")) = "wife's", "Match position wrong"
        End If
    Next
    Check found, "No finding object"
    dismissal = RestrictedDismiss(id, "RC-FAMILY", "OfficialComments", "wife's", "Reviewed with the senior rater")
    Check FindingCount(id, "RC-FAMILY") = 0 And FindingCount(other, "RC-FAMILY") = 1, "Dismissal not limited to its report"
    Check DCount("*", "OperationJournal", "OperationKind='Dismiss content finding'") = 1, "Dismissal not recorded"
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "OfficialComments", "Despite his wife's illness and his husband duties, he never missed a watch."
    Call CommitSavedFields(id, CurrentRevision(id), changes, "New wording")
    Check FindingCount(id, "RC-FAMILY") = 1, "New words were hidden by an earlier dismissal"
    Check InStr(RestrictedListRows(), "D:" & dismissal) > 0, "Dismissal not listed"
    RestrictedRestoreDismissal dismissal
    Check FindingCount(id, "RC-FAMILY") = 2, "Restored dismissal did not bring the finding back"

    ' Allow a phrase for one rule on every report; removing it brings findings back.
    exception = RestrictedAllow("wife", "RC-FAMILY", "Command decision")
    Check FindingCount(other, "RC-FAMILY") = 0 And Flags("His wife deployed with him.") = "", "Allowed phrase still flagged"
    Check Flags("His wife's surgery went well.") = "RC-MEDICAL", "Allowed phrase hid another rule"
    Check BadAllow("wife", "RC-FAMILY") And BadAllow("", "") And BadAllow("x", "RC-NOT-A-RULE"), "Invalid or duplicate allowed phrase accepted"
    RestrictedRemoveException exception
    Check FindingCount(other, "RC-FAMILY") = 1, "Removed phrase still allowed"
    exception = RestrictedAllow("fleet week", "", "")
    Check Flags("His wife attended fleet week.") = "RC-FAMILY", "All-rule phrase hid an unrelated word"
    RestrictedRemoveException exception

    ' A full pass over 5,000 characters is fast (cached passes are near free).
    text = ""
    Do While Len(text) < 5000: text = text & "Led a team of twelve Sailors through a demanding inspection cycle, earning zero discrepancies and a Navy and Marine Corps Achievement Medal. ": Loop
    RestrictedInvalidate
    started = Timer
    For i = 1 To 5: Call RestrictedMatches("EVAL", "OfficialComments", text & i, 0): Next
    elapsed = (Timer - started) * 1000 / 5
    Check elapsed < 100, "Content check took " & Format$(elapsed, "0.0") & " ms for 5,000 characters"
    TestRestrictedContent = "PASS: " & checks & " restricted-content rule, false-positive, dismissal, allowed-phrase and timing assertions; " & Format$(elapsed, "0.0") & " ms per 5,000-character pass": Exit Function
Failed:
    TestRestrictedContent = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
Private Function BadAllow(ByVal phrase As String, ByVal rule As String) As Boolean
    On Error GoTo Refused
    Call RestrictedAllow(phrase, rule, "")
    Exit Function
Refused:
    BadAllow = True
End Function
