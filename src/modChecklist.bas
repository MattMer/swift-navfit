Attribute VB_Name = "modChecklist"
Option Compare Database
Option Explicit
Private Function HasContent(ByVal value As Variant) As Boolean
    Dim text As String
    text = Nz(value, "")
    text = Replace(Replace(Replace(text, vbCr, ""), vbLf, ""), vbTab, "")
    HasContent = Len(Trim$(text)) > 0
End Function
Private Sub Finding(ByVal key As String, ByVal severity As String, ByVal message As String)
    'Only fixed application text enters the value list, never source/user values.
    Forms!frmChecklist!lstFindings.AddItem key & ";" & severity & ";" & message
End Sub
Public Function OpenDraftChecklist(ByVal reportID As Long) As Boolean
    If reportID = 0 Or DCount("*", "Report", "ReportID=" & reportID) = 0 Then Err.Raise vbObjectError + 1000, , "Save or select a draft first."
    DoCmd.OpenForm "frmChecklist"
    Forms!frmChecklist!txtReportID = reportID
    Call RefreshDraftChecklist()
    Forms!frmChecklist.SetFocus
    OpenDraftChecklist = True
End Function
Public Function ChecklistFromEditor() As Boolean
    On Error GoTo Failed
    ChecklistFromEditor = OpenDraftChecklist(Nz(Forms!frmEditor!txtReportID, 0))
    Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Saved draft checklist"
End Function
Public Function ChecklistFromLibrary() As Boolean
    On Error GoTo Failed
    ChecklistFromLibrary = OpenDraftChecklist(Nz(Forms!frmLibrary!lstReports, 0))
    Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Saved draft checklist"
End Function
Public Function RefreshDraftChecklist() As Boolean
    Dim r As DAO.Recordset, id As Long, missing As Long, errors As Long, context As String, key As Variant, hasNOB As Boolean
    id = Nz(Forms!frmChecklist!txtReportID, 0)
    Set r = CurrentDb.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 1001, , "Draft no longer exists."
    Forms!frmChecklist!lstFindings.RowSourceType = "Value List"
    Forms!frmChecklist!lstFindings.RowSource = ""
    If Not HasContent(r!Title) Then Finding "TITLE", "Missing input", "Add a local draft title to identify this report.": missing = missing + 1
    If Not HasContent(r!MemberName) Then Finding "MEMBER", "Missing input", "Add a local member display name. It stays separate from the official full name.": missing = missing + 1
    For Each key In AdministrationKeys()
        If Not HasContent(r.Fields(key).Value) Then
            Finding CStr(key), "Missing input", "Review " & AdministrationLabel(key) & ". Presence only, not a Navy policy requirement.": missing = missing + 1
        End If
    Next
    If Not HasContent(r!PromotionStatus) Then Finding "PromotionStatus", "Missing input", "Review promotion status. Presence only; eligibility is not assessed.": missing = missing + 1
    For Each key In ReportContextKeys()
        If Not HasContent(r.Fields(key).Value) Then Finding CStr(key), "Missing input", "Review " & ReportContextLabel(key) & ". Presence only; policy is not assessed.": missing = missing + 1
    Next
    For Each key In ReportingSeniorKeys()
        If Not HasContent(r.Fields(key).Value) Then Finding CStr(key), "Missing input", "Review " & ReportingSeniorLabel(key) & ". Presence only; policy is not assessed.": missing = missing + 1
    Next
    For Each key In AssessmentKeys()
        If Not HasContent(r.Fields(key).Value) Then Finding CStr(key), "Missing input", "Review " & AssessmentLabel(key) & ". Explicit selection only; policy applicability is not assessed.": missing = missing + 1
    Next
    For Each key In TraitKeys()
        If Not HasContent(r.Fields(key).Value) Then
            Finding CStr(key), "Missing input", "Review trait block " & Mid$(key, 6) & ". Blank is distinct from NOB; incomplete saving remains available.": missing = missing + 1
        ElseIf r.Fields(key).Value = "NOB" Then
            hasNOB = True
        End If
    Next
    If hasNOB Then Finding "TRAIT_NOB", "Review", "Individual NOB marks do not establish whole-report observation status or policy compliance."
    If IsNull(r!PeriodStart) Then Finding "START", "Missing input", "Enter the period start date.": missing = missing + 1
    If IsNull(r!PeriodEnd) Then Finding "END", "Missing input", "Enter the period end date.": missing = missing + 1
    If Not IsNull(r!PeriodStart) And Not IsNull(r!PeriodEnd) Then
        If r!PeriodEnd < r!PeriodStart Then Finding "DATE_ORDER", "Draft error", "Period end precedes period start. Review the dates.": errors = errors + 1
    End If
    If Not HasContent(r!OfficialComments) Then Finding "NARRATIVE", "Missing input", "The official narrative is blank. General draft notes do not fill it.": missing = missing + 1
    If DCount("*", "ImportReceipt", "ReportID=" & id) > 0 Then Finding "PARTIAL_IMPORT", "Needs mapping", "This draft imported a narrative only. Other source fields remain in provenance and are unmapped."
    If DCount("*", "RecoveryDraft", "ReportID=" & id) > 0 Then Finding "RECOVERY", "Review", "A pending-edit checkpoint exists. It is excluded from this saved-draft checklist."
    If r!IsArchived Then Finding "ARCHIVED", "Review", "This draft is archived. Recover it before editing."
    Finding "POLICY", "Not available", "Navy policy validation and official field completeness checks are not implemented."
    Finding "EXPORT", "Not available", "NAVFIT export readiness is not assessed. Export is not implemented."
    context = "Report " & id & " - saved revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & Nz(r!CurrentRevisionID, 0)), 0)
    Forms!frmChecklist!lblContext.Caption = context
    Forms!frmChecklist!lblSummary.Caption = missing & " missing local inputs; " & errors & " draft date errors. Saved values only; unsaved editor entries are excluded."
    Forms!frmChecklist!txtGuidance = "Select a checklist item for guidance. Draft saving remains available even when inputs are incomplete."
    r.Close
    RefreshDraftChecklist = True
End Function
Public Function ChecklistGuidance() As Boolean
    Dim key As String, guidance As String
    key = Nz(Forms!frmChecklist!lstFindings, "")
    Select Case key
        Case "PromotionStatus": guidance = "Choose REGULAR, FROCKED, SELECTED or SPOT in the editor, or leave blank. This is separate from promotion recommendation. No inference, eligibility or Navy policy validation is performed."
        Case "TITLE": guidance = "In the editor, fill Draft title, then Save draft. This is a local organizational field."
        Case "MEMBER": guidance = "In the editor, fill Member name, then Save draft. This display name does not validate official name components, DoD ID or SSN."
        Case "START", "END", "DATE_ORDER": guidance = "Review Period start and Period end in the editor, then Save draft. This checklist checks presence and chronological order only; occasions, continuity and policy applicability are not evaluated."
        Case "NARRATIVE": guidance = "Scroll below revision history to the official narrative section. EVAL uses block 43, CHIEFEVAL block 40 and FITREP block 41. Draft notes stay separate. A blank narrative can be saved; this is a completeness prompt, not a policy finding."
        Case "PARTIAL_IMPORT": guidance = "Narrative import creates a partial draft. Filling local title, member and dates does not complete the unmapped official administration, traits, recommendations or folder relationships. Complete import mapping is pending."
        Case "RECOVERY": guidance = "Use Home > Recover pending edits to inspect the checkpoint. Recovery is separate from committed revisions. If the saved report changed, recover as a new draft."
        Case "ARCHIVED": guidance = "Use the library's Archived or All filter, select the report and choose Recover."
        Case "POLICY": guidance = "This checklist is not Navy policy validation. The versioned, source-linked rule catalogue for official fields, grouping, recommendations, dates and report conditions remains pending."
        Case "EXPORT": guidance = "A complete local checklist does not mark this report approved or ready for NAVFIT export. Export and full round-trip verification remain pending."
        Case "TRAIT_NOB": guidance = "Trait NOB is a distinct listed selection, not a zero score or blank. Whole-report NOB, partial observation, averages and related narrative requirements are not assessed in this build."
        Case Else
            If IsTraitKey(key) Then
                guidance = "Use the form-specific trait section in the editor. Select a mark or leave it blank, then Save draft. Select a trait to read the NAVFIT reference descriptions. Recommendations, averages and observation rules are not yet validated."
            ElseIf Len(ReportContextLabel(key)) > 0 Then
                guidance = "Open the relevant report-context panel in the editor, apply pending values, then Save draft. Yes, No and blank are distinct. Occasion/category combinations and duty narrative rules are not policy-validated."
            ElseIf Len(ReportingSeniorLabel(key)) > 0 Then
                guidance = "Open Reporting senior details in the editor, apply pending values, then Save draft. These fields are separate from the local reviewer profile. Storage limits do not validate codes or reporting authority."
            ElseIf Len(AssessmentLabel(key)) > 0 Then
                guidance = "Choose observation status and recommendation explicitly near the top of the editor. Individual trait NOB and averages do not set them. Incomplete saves remain available; eligibility and policy are not validated."
            ElseIf Len(AdministrationLabel(key)) > 0 Then
                guidance = "Review the administrative fields near the top of the editor, then Save draft. Storage limits match the inspected NAVFIT schema. Values stay separate from local display fields. Codes, format, eligibility and automatic NAVFIT mapping are not validated. Incomplete drafts remain saveable."
            Else
                guidance = "Select an item."
            End If
    End Select
    Forms!frmChecklist!txtGuidance = guidance
    ChecklistGuidance = True
End Function
Public Function CloseDraftChecklist() As Boolean
    DoCmd.Close acForm, "frmChecklist", acSaveNo
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Forms!frmEditor.SetFocus Else Call OpenScreen("frmLibrary")
    CloseDraftChecklist = True
End Function
