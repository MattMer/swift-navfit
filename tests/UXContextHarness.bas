Attribute VB_Name = "UXContextHarness"
Option Compare Database
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 7000, , message
    checks = checks + 1
End Sub
Public Function TestUXContext() As String
    Dim id As Long, groupID As Long, values As Object, key As Variant, text As String, revision As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC CONTEXT", "Synthetic Tester"
    id = CommitDraft(0, "EVAL", "Synthetic context", "Synthetic member", #1/1/2026#, #9/30/2026#, "", "Test")
    Set values = CreateObject("Scripting.Dictionary")
    For Each key In TraitKeys(): values.Add CStr(key), "4.0": Next
    values.Add "ObservationStatus", "Observed"
    Call CommitSavedFields(id, CurrentRevision(id), values, "Test marks")
    groupID = CreateSummaryGroup("Synthetic context group", "EVAL", "E5"): AddGroupMember groupID, id, "E5"
    text = UXGroupCalculationText(groupID)
    Check InStr(text, "28.00 trait sum / 7 graded traits = 4.00") > 0, "Group formula/count wrong"
    Check InStr(text, "pending report/comparison edits are excluded") > 0, "Saved source missing"
    values("Trait33") = "NOB": values("Trait34") = ""
    text = UXTraitContext(values)
    Check InStr(text, "5 graded traits") > 0 And InStr(text, "1 NOB, 1 blank") > 0, "Partial exclusions wrong"
    values("ObservationStatus") = "NOB": Check InStr(UXTraitContext(values), "all seven traits excluded") > 0, "Whole NOB exclusion wrong"
    Check ChecklistField("NARRATIVE") = "OfficialComments", "Narrative target wrong"
    Check UXFieldSection(ChecklistField("DATE_ORDER")) = 0, "Date target wrong"
    Check UXFieldSection(ChecklistField("TRAIT_NOB")) = 4, "Traits target wrong"
    Check OpenDraftChecklist(id), "Checklist failed"
    Forms!frmChecklist!lstFindings = "NARRATIVE": Check ChecklistGuidance(), "Guidance failed"
    Check Forms!frmChecklist!cmdOpenSection.Enabled, "Field action disabled"
    revision = CurrentRevision(id): Check ChecklistOpenSection(), "Section jump failed"
    Check Forms!frmEditor!uxSection.Caption = "Recommendations && comments", "Narrative section wrong"
    Check CurrentRevision(id) = revision, "Jump created revision"
    Check Forms!frmEditor!cmdChecklist.OnClick = "=ChecklistFromEditor()", "Editor checklist action misrouted"
    Forms!frmEditor!txtNotes = "Pending notes survive checklist navigation"
    Forms!frmChecklist!lstFindings = "START": Call ChecklistGuidance
    Check ChecklistOpenSection(), "Second section jump failed"
    Check Forms!frmEditor!txtNotes = "Pending notes survive checklist navigation", "Checklist navigation lost pending edits"
    Check CurrentRevision(id) = revision, "Pending navigation committed revision"
    Check InStr(UXFieldCaption(ChecklistField("NARRATIVE"), "CHIEFEVAL"), "40.") = 1, "Chief narrative block wrong"
    Check InStr(UXFieldCaption(ChecklistField("NARRATIVE"), "FITREP"), "41.") = 1, "FITREP narrative block wrong"
    Forms!frmChecklist!lstFindings = "POLICY": Call ChecklistGuidance
    Check Not Forms!frmChecklist!cmdOpenSection.Enabled, "Nonfield action enabled"
    Check InStr(Forms!frmChecklist!txtGuidance, "effective report-period range") > 0, "Limited check scope missing"
    DoCmd.OpenForm "frmUXGroupRSCA", , , , , , CStr(groupID)
    Check InStr(Forms!frmUXGroupRSCA!txtCalculationContext, "7 graded traits") > 0, "Initial calculation context missing"
    Check InStr(Forms!frmUXGroupRSCA!lblStatus.Caption, "Choose a reporting senior") > 0, "Unavailable reason missing"
    TestUXContext = "PASS: " & checks & " calculation context/checklist assertions": Exit Function
Failed:
    TestUXContext = "FAIL: " & Err.Description
End Function
