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
    Dim field As String, kind As String
    field = ChecklistField(key): kind = Nz(DLookup("FormKind", "Report", "ReportID=" & Nz(Forms!frmChecklist!txtReportID, 0)), "EVAL")
    If Len(field) > 0 Then
        If UXFieldLegacy(field) Or Not UXFieldApplies(field, kind) Then Exit Sub
        message = UXFieldCaption(field, kind) & ": " & message
    End If
    Forms!frmChecklist!lstFindings.AddItem key & ";" & severity & ";" & Replace(message, ";", ",")
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
    Dim r As DAO.Recordset, id As Long, missing As Long, errors As Long, context As String, key As Variant, hasNOB As Boolean, item As Long
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
    For Each key In PartyNamesKeys()
        If Not HasContent(r.Fields(key).Value) Then Finding CStr(key), "Missing input", "Review " & PartyNamesLabel(key) & ". Presence only; applicability is not assessed.": missing = missing + 1
    Next
    For Each key In ReportPeopleKeys()
        If Not HasContent(r.Fields(key).Value) Then Finding CStr(key), "Missing input", "Review " & ReportPeopleLabel(key) & ". Presence only; applicability is not assessed.": missing = missing + 1
    Next
    For Each key In DecisionDetailsKeys()
        If Not HasContent(r.Fields(key).Value) Then Finding CStr(key), "Missing input", "Review " & DecisionDetailsLabel(key) & ". Presence only; applicability is not assessed.": missing = missing + 1
    Next
    For Each key In ServiceDetailsKeys()
        If Not HasContent(r.Fields(key).Value) Then Finding CStr(key), "Missing input", "Review " & ServiceDetailsLabel(key) & ". Presence only; policy is not assessed.": missing = missing + 1
    Next
    For Each key In MemberDetailsKeys()
        If Not HasContent(r.Fields(key).Value) Then Finding CStr(key), "Missing input", "Review " & MemberDetailsLabel(key) & ". Presence only; policy is not assessed.": missing = missing + 1
    Next
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
    Finding "POLICY", "Limited Navy checks", "Source-linked draft checks are available in Review. They cover selected rules only and do not certify compliance."
    Finding "EXPORT", "Separate workflow", "Import/export previews and compatibility checks are separate from this completeness checklist."
    context = UXMemberLabel(id) & " - saved revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & Nz(r!CurrentRevisionID, 0)), 0)
    Forms!frmChecklist!lblContext.Caption = context
    missing = 0
    For item = 0 To Forms!frmChecklist!lstFindings.ListCount - 1
        If Forms!frmChecklist!lstFindings.Column(1, item) = "Missing input" Then missing = missing + 1
    Next
    Forms!frmChecklist!lblSummary.Caption = missing & " missing local inputs; " & errors & " draft date errors. Saved values only; unsaved editor entries are excluded."
    Forms!frmChecklist!cmdOpenSection.Enabled = False
    Forms!frmChecklist!txtGuidance = "Select a checklist item for guidance. Draft saving remains available even when inputs are incomplete."
    r.Close
    RefreshDraftChecklist = True
End Function
Public Function ChecklistField(ByVal key As String) As String
    Select Case key
        Case "TITLE": ChecklistField = "Title"
        Case "MEMBER": ChecklistField = "MemberName"
        Case "START", "DATE_ORDER": ChecklistField = "PeriodStart"
        Case "END": ChecklistField = "PeriodEnd"
        Case "NARRATIVE": ChecklistField = "OfficialComments"
        Case "TRAIT_NOB": ChecklistField = "Trait33"
        Case "POLICY", "EXPORT", "PARTIAL_IMPORT", "RECOVERY", "ARCHIVED": ChecklistField = ""
        Case Else: ChecklistField = key
    End Select
End Function
Public Function ChecklistGuidance() As Boolean
    Dim key As String, field As String, kind As String, guidance As String
    key = Nz(Forms!frmChecklist!lstFindings, ""): field = ChecklistField(key)
    kind = Nz(DLookup("FormKind", "Report", "ReportID=" & Forms!frmChecklist!txtReportID), "EVAL")
    If Len(field) > 0 Then
        guidance = UXFieldCaption(field, kind) & vbCrLf & "Choose Open report section to review this field. Incomplete draft information can be saved. This checklist reads saved values only; pending edits and recovery checkpoints are separate. Name parts never replace separately stored full names."
    Else
        Select Case key
            Case "POLICY": guidance = "Review > Navy checks uses a versioned source-linked rule pack for selected checks. Check the pack's effective report-period range and saved revision. Needs review differs from an invalid entry; a passing limited check does not establish full Navy compliance."
            Case "EXPORT", "PARTIAL_IMPORT": guidance = "Use Reports > Import or the report's export workflow. Inspect the mapping scope and preview; this checklist does not establish export readiness."
            Case "RECOVERY": guidance = "Overview > Recover pending edits. Inspect the checkpoint before recovery; saved revisions remain separate."
            Case "ARCHIVED": guidance = "Reports > Archived or All, select this report and choose Recover. Local approval also locks edits until reopened through Review."
            Case Else: guidance = "Select a checklist item."
        End Select
    End If
    Forms!frmChecklist!txtGuidance = guidance
    Forms!frmChecklist!cmdOpenSection.Enabled = Len(field) > 0
    ChecklistGuidance = True
End Function
Public Function ChecklistOpenSection() As Boolean
    Dim id As Long, field As String
    On Error GoTo Failed
    id = Forms!frmChecklist!txtReportID: field = ChecklistField(Nz(Forms!frmChecklist!lstFindings, ""))
    If Len(field) = 0 Then Exit Function
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Nz(Forms!frmEditor!txtReportID, 0) <> id Then
            If Not UXLeaveEditor() Then Exit Function
        End If
    End If
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Call UXSelectSection(UXFieldSection(field))
    Forms!frmEditor.SetFocus
    ChecklistOpenSection = True: Exit Function
Failed:
    Forms!frmChecklist!txtGuidance = Err.Description
End Function
Public Function CloseDraftChecklist() As Boolean
    DoCmd.Close acForm, "frmChecklist", acSaveNo
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Forms!frmEditor.SetFocus Else Call OpenScreen("frmLibrary")
    CloseDraftChecklist = True
End Function
