Attribute VB_Name = "modValidation"
Option Compare Binary
Option Explicit
Public Const NAVY_RULE_SOURCE_HASH As String = "5E71CA59D66C509640B48845960D236670ABFB5D57AE76435A7C7983F0DC9058"
Public Function PromotionCapacity(ByVal paygrade As String, ByVal observedGroupSize As Long, ByVal isLDO As Boolean, Optional ByVal earlyPromotesUsed As Long = -1) As Object
    Dim result As Object, percent As Long, ep As Long, combined As Long
    If observedGroupSize < 1 Or observedGroupSize > 1000000 Then Err.Raise vbObjectError + 4510, , "Enter a positive, verified observed group size."
    Select Case paygrade
        Case "E1", "E2", "E3", "E4", "W1", "W2": percent = 100
        Case "E5", "E6", "O3": percent = 60
        Case "E7", "E8", "E9", "W3", "W4", "W5", "O4": percent = 50
        Case "O5", "O6": percent = 40
        Case "O1", "O2": If isLDO Then percent = 100 Else percent = 0
        Case Else: Err.Raise vbObjectError + 4511, , "Unsupported paygrade requires policy review."
    End Select
    If percent > 0 Then
        ep = (observedGroupSize * 20 + 99) \ 100
        combined = (observedGroupSize * percent + 99) \ 100
        If observedGroupSize = 2 Then combined = 2
    End If
    If earlyPromotesUsed = -1 Then earlyPromotesUsed = ep
    If earlyPromotesUsed < 0 Or earlyPromotesUsed > ep Then Err.Raise vbObjectError + 4512, , "Early Promote count exceeds its capacity."
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "EarlyPromoteMaximum", ep: result.Add "CombinedMaximum", combined: result.Add "MustPromoteMaximum", combined - earlyPromotesUsed
    result.Add "Source", "BUPERSINST 1610.10H CH-2, enclosure 2, 1-17 and table 1-2; observed, complete eligible groups only"
    result.Add "SourceHash", NAVY_RULE_SOURCE_HASH
    Set PromotionCapacity = result
End Function
Public Function InstallNavyDraftChecks() As Long
    Dim p As DAO.Recordset, d As DAO.Recordset, w As DAO.Workspace, id As Long, rule As Variant, failure As Long, message As String
    id = Nz(MajorScalar("SELECT PackID FROM RulePack WHERE VersionName='1610.10H CH-2 / draft-checks-v1'"), 0)
    If id > 0 Then InstallNavyDraftChecks = id: Exit Function
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set p = CurrentDb.OpenRecordset("RulePack", dbOpenDynaset)
    p.AddNew: p!VersionName = "1610.10H CH-2 / draft-checks-v1": p!SourceTitle = "BUPERSINST 1610.10H CH-2, 26 May 2026"
    p!SourceURL = "https://www.mynavyhr.navy.mil/References/BUPERS-Instructions/": p!SourceSHA256 = NAVY_RULE_SOURCE_HASH
    p!EffectiveFrom = #5/26/2026#: p!EffectiveTo = #10/3/2026#: p.Update: p.Bookmark = p.LastModified: id = p!PackID: p.Close
    Set d = CurrentDb.OpenRecordset("RuleDefinition", dbOpenDynaset)
    For Each rule In Array(Array("NAME-PRESENT", "OfficialName", "Required", "Enclosure 2, 1-1, block 1", "Official name requires review when absent."), _
        Array("RATE-PRESENT", "RankRate", "Required", "Enclosure 2, 1-1, block 2", "Grade/rate worn at the report ending date is required."), _
        Array("UIC-PRESENT", "UnitUIC", "Required", "Enclosure 2, chapter 1, block 6", "Unit UIC requires review when absent."), _
        Array("TRAIT-RECOMMENDATION", "PromotionRecommendation", "TraitPromotion", "Enclosure 2, 1-16", "Trait grades constrain promotion recommendations; incomplete marks require review."), _
        Array("JUNIOR-OFFICER", "PromotionRecommendation", "JuniorOfficer", "Enclosure 2, 1-16, note 2", "Non-LDO ENS/LTJG recommendations cannot exceed Promotable."))
        d.AddNew: d!PackID = id: d!RuleCode = rule(0): d!FieldKey = rule(1): d!CheckKind = rule(2): d!SourceLocator = rule(3): d!Description = rule(4): d!Severity = "Review"
        If rule(0) = "JUNIOR-OFFICER" Then d!FormKind = "FITREP"
        d.Update
    Next
    d.Close: w.CommitTrans: InstallNavyDraftChecks = id: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not p Is Nothing Then p.Close
    If Not d Is Nothing Then d.Close
    On Error GoTo 0
    Err.Raise failure, "InstallNavyDraftChecks", message
End Function
Private Function TraitPromotionFinding(ByVal fields As Object) As String
    Dim key As Variant, value As String, recommendation As String, incomplete As Boolean
    recommendation = Nz(fields("PromotionRecommendation"), "")
    If recommendation = "" Then TraitPromotionFinding = "NeedsReview": Exit Function
    TraitPromotionFinding = "Pass"
    For Each key In TraitKeys()
        value = Nz(fields(CStr(key)), "")
        If value = "" Or value = "NOB" Then incomplete = True
        If value = "1.0" And (recommendation = "Promotable" Or recommendation = "Must Promote" Or recommendation = "Early Promote") Then TraitPromotionFinding = "Fail": Exit Function
        If value = "2.0" And (recommendation = "Must Promote" Or recommendation = "Early Promote") Then TraitPromotionFinding = "Fail": Exit Function
    Next
    If incomplete Then TraitPromotionFinding = "NeedsReview"
End Function
' Historical rule definitions are immutable inputs; a run always names a pack and saved revision.
Public Function RunRulePack(ByVal reportID As Long, ByVal packID As Long) As Long
    Dim p As DAO.Recordset, d As DAO.Recordset, r As DAO.Recordset, f As DAO.Recordset, fields As Object, id As Long, revision As Long, state As String, value As Variant
    Dim w As DAO.Workspace, failure As Long, message As String
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    revision = CurrentRevision(reportID): Set fields = SavedFields(reportID)
    Set p = CurrentDb.OpenRecordset("SELECT * FROM RulePack WHERE PackID=" & packID, dbOpenSnapshot)
    If p.EOF Then Err.Raise vbObjectError + 4500, , "Rule pack does not exist."
    If Len(Nz(p!SourceSHA256, "")) <> 64 Or Len(Nz(p!SourceURL, "")) = 0 Then Err.Raise vbObjectError + 4501, , "Rule pack requires a source fingerprint and URL."
    Set r = CurrentDb.OpenRecordset("ValidationRun", dbOpenDynaset)
    r.AddNew: r!ReportID = reportID: r!RevisionID = revision: r!PackID = packID: r!RunAt = Now(): r.Update: r.Bookmark = r.LastModified: id = r!ValidationRunID: r.Close
    Set d = CurrentDb.OpenRecordset("SELECT * FROM RuleDefinition WHERE PackID=" & packID & " ORDER BY RuleID", dbOpenSnapshot)
    If d.EOF Then Err.Raise vbObjectError + 4502, , "Rule pack is empty."
    Set f = CurrentDb.OpenRecordset("ValidationFinding", dbOpenDynaset)
    Do While Not d.EOF
        If Nz(d!FormKind, "") = "" Or d!FormKind = fields("FormKind") Then
            state = "NeedsReview"
            If Not IsNull(fields("PeriodEnd")) And Not IsNull(p!EffectiveFrom) And Not IsNull(p!EffectiveTo) Then
                If fields("PeriodEnd") >= p!EffectiveFrom And fields("PeriodEnd") <= p!EffectiveTo And fields.Exists(Nz(d!FieldKey, "")) And Len(Nz(d!SourceLocator, "")) > 0 Then
                    value = fields(CStr(d!FieldKey))
                    Select Case Nz(d!CheckKind, "")
                        Case "Required"
                            If Len(Trim$(Nz(value, ""))) > 0 Then state = "Pass" Else state = "Fail"
                        Case "Equals"
                            If SameFieldValue(value, d!ExpectedValue) Then state = "Pass" Else state = "Fail"
                        Case "TraitPromotion": state = TraitPromotionFinding(fields)
                        Case "JuniorOfficer"
                            If fields("FormKind") = "FITREP" And Len(Nz(fields("RankRate"), "")) > 0 And Len(Nz(fields("Designator"), "")) = 4 Then
                                state = "Pass"
                                If fields("RankRate") = "ENS" Or fields("RankRate") = "LTJG" Then
                                    If Left$(fields("Designator"), 1) <> "6" And (value = "Early Promote" Or value = "Must Promote") Then state = "Fail"
                                End If
                            End If
                    End Select
                End If
            End If
            f.AddNew: f!ValidationRunID = id: f!RuleID = d!RuleID: f!FindingState = state: f!Message = Nz(d!Description, "") & " [" & Nz(d!SourceLocator, "") & "]": f.Update
        End If
        d.MoveNext
    Loop
    f.Close: d.Close: p.Close: w.CommitTrans: RunRulePack = id: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    w.Rollback
    If Not p Is Nothing Then p.Close
    If Not d Is Nothing Then d.Close
    If Not r Is Nothing Then r.Close
    If Not f Is Nothing Then f.Close
    On Error GoTo 0
    Err.Raise failure, "RunRulePack", message
End Function
