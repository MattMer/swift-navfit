Attribute VB_Name = "modRevisionCompare"
Option Compare Database
Option Explicit

Private Function SavedValue(ByVal revisionID As Long, ByVal key As String, ByRef recorded As Boolean) As Variant
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='" & key & "'", dbOpenSnapshot)
    recorded = Not r.EOF
    If Not recorded Then
        If key <> "OfficialComments" And Len(AdministrationLabel(key)) = 0 And Not IsTraitKey(key) And Len(AssessmentLabel(key)) = 0 And key <> "PromotionStatus" And Len(ReportingSeniorLabel(key)) = 0 And Len(ReportContextLabel(key)) = 0 And Len(MemberDetailsLabel(key)) = 0 And Len(ServiceDetailsLabel(key)) = 0 And Len(DecisionDetailsLabel(key)) = 0 And Len(ReportPeopleLabel(key)) = 0 And Len(PartyNamesLabel(key)) = 0 Then Err.Raise vbObjectError + 800, , "Revision snapshot is incomplete."
        If Len(AdministrationLabel(key)) > 0 And RevisionSnapshotSchema(revisionID) >= 9 Then Err.Raise vbObjectError + 800, , "Administrative snapshot is incomplete."
        If key = "PromotionStatus" And RevisionSnapshotSchema(revisionID) >= 12 Then Err.Raise vbObjectError + 800, , "Promotion status snapshot is incomplete."
        If Len(ReportingSeniorLabel(key)) > 0 And RevisionSnapshotSchema(revisionID) >= 13 Then Err.Raise vbObjectError + 800, , "Reporting senior snapshot is incomplete."
        If Len(ReportContextLabel(key)) > 0 And RevisionSnapshotSchema(revisionID) >= 14 Then Err.Raise vbObjectError + 800, , "Report context snapshot is incomplete."
        If Len(MemberDetailsLabel(key)) > 0 And RevisionSnapshotSchema(revisionID) >= 15 Then Err.Raise vbObjectError + 800, , "Member details snapshot is incomplete."
        If Len(ServiceDetailsLabel(key)) > 0 And RevisionSnapshotSchema(revisionID) >= 16 Then Err.Raise vbObjectError + 800, , "Service details snapshot is incomplete."
        If Len(DecisionDetailsLabel(key)) > 0 And RevisionSnapshotSchema(revisionID) >= 17 Then Err.Raise vbObjectError + 800, , "Recommendation and statement snapshot is incomplete."
        If Len(ReportPeopleLabel(key)) > 0 And RevisionSnapshotSchema(revisionID) >= 18 Then Err.Raise vbObjectError + 800, , "Reporting contact and rater snapshot is incomplete."
        If Len(PartyNamesLabel(key)) > 0 And RevisionSnapshotSchema(revisionID) >= 19 Then Err.Raise vbObjectError + 800, , "Structured names snapshot is incomplete."
        If IsTraitKey(key) And RevisionSnapshotSchema(revisionID) >= 10 Then Err.Raise vbObjectError + 800, , "Trait snapshot is incomplete."
        If Len(AssessmentLabel(key)) > 0 And RevisionSnapshotSchema(revisionID) >= 11 Then Err.Raise vbObjectError + 800, , "Assessment snapshot is incomplete."
        SavedValue = Null
    ElseIf r!IsNullValue Then
        SavedValue = Null
    Else
        If IsNull(r!ValueText) Then Err.Raise vbObjectError + 801, , "Revision snapshot has an inconsistent null flag."
        SavedValue = r!ValueText
    End If
    r.Close
End Function

Public Function RevisionFieldStatus(ByVal older As Long, ByVal newer As Long, ByVal key As String) As String
    Dim a As Variant, b As Variant, ar As Boolean, br As Boolean
    a = SavedValue(older, key, ar): b = SavedValue(newer, key, br)
    If Not ar And Not br Then
        RevisionFieldStatus = "Not recorded"
    ElseIf ar <> br Or IsNull(a) <> IsNull(b) Then
        RevisionFieldStatus = "Changed"
    ElseIf StrComp(Nz(a, ""), Nz(b, ""), vbBinaryCompare) <> 0 Then
        RevisionFieldStatus = "Changed"
    Else
        RevisionFieldStatus = "Unchanged"
    End If
End Function

Private Function DisplayValue(ByVal revisionID As Long, ByVal key As String) As String
    Dim value As Variant, recorded As Boolean
    value = SavedValue(revisionID, key, recorded)
    If Not recorded Then
        DisplayValue = "[Not recorded in this revision]"
    ElseIf IsNull(value) Then
        DisplayValue = "[NULL]"
    ElseIf Len(value) = 0 Then
        DisplayValue = "[Empty text]"
    Else
        DisplayValue = CStr(value)
    End If
End Function

Public Function OpenRevisionComparison(ByVal reportID As Long, ByVal older As Long, Optional ByVal newer As Long = 0) As Boolean
    Dim r As DAO.Recordset, recorded As Boolean, ignored As Variant, key As Variant
    If newer = 0 Then newer = Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & reportID), 0)
    If older = 0 Or newer = 0 Then Err.Raise vbObjectError + 802, , "Select a saved revision first."
    If Nz(DLookup("ReportID", "ReportRevision", "RevisionID=" & older), 0) <> reportID Or Nz(DLookup("ReportID", "ReportRevision", "RevisionID=" & newer), 0) <> reportID Then Err.Raise vbObjectError + 803, , "Both revisions must belong to this report."
    'Validate before displaying a new comparison; no writes or editor saves occur.
    For Each key In DraftFieldKeys()
        ignored = SavedValue(older, CStr(key), recorded): ignored = SavedValue(newer, CStr(key), recorded)
    Next
    DoCmd.OpenForm "frmRevisionCompare"
    With Forms!frmRevisionCompare
        !txtOlder = older: !txtNewer = newer: !chkChangedOnly = True
        !lblOlder.Caption = "Selected saved revision " & DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & older)
        !lblNewer.Caption = "Compared saved revision " & DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & newer)
    End With
    Call RefreshComparisonFields()
    Forms!frmRevisionCompare.SetFocus
    OpenRevisionComparison = True
End Function

Public Function CompareSelectedRevision() As Boolean
    On Error GoTo Failed
    CompareSelectedRevision = OpenRevisionComparison(Nz(Forms!frmEditor!txtReportID, 0), Nz(Forms!frmEditor!lstRevisions, 0))
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Compare revisions"
End Function

Public Function RefreshComparisonFields() As Boolean
    Dim f As Form, key As Variant, status As String, changed As Long, shown As Long
    Set f = Forms!frmRevisionCompare
    f!lstCompareFields.RowSourceType = "Value List": f!lstCompareFields.RowSource = ""
    For Each key In DraftFieldKeys()
        status = RevisionFieldStatus(CLng(f!txtOlder), CLng(f!txtNewer), CStr(key))
        If status = "Changed" Then changed = changed + 1
        If Not f!chkChangedOnly Or status = "Changed" Then
            f!lstCompareFields.AddItem CStr(key) & ";" & ComparisonFieldLabel(CStr(key)) & ";" & status
            shown = shown + 1
        End If
    Next
    f!lblSummary.Caption = changed & " changed fields. Saved snapshots only; unsaved editor entries are excluded."
    f!txtOldValue = "": f!txtNewValue = ""
    If shown > 0 Then
        f!lstCompareFields = f!lstCompareFields.ItemData(0)
        Call LoadComparisonValues()
    Else
        f!txtOldValue = "No changed fields. Uncheck Changed fields only to inspect all fields."
    End If
    RefreshComparisonFields = True
End Function

Private Function ComparisonFieldLabel(ByVal key As String) As String
    Select Case key
        Case "FormKind": ComparisonFieldLabel = "Report type"
        Case "Title": ComparisonFieldLabel = "Draft title"
        Case "MemberName": ComparisonFieldLabel = "Member display name"
        Case "PeriodStart": ComparisonFieldLabel = "Period start"
        Case "PeriodEnd": ComparisonFieldLabel = "Period end"
        Case "DraftNotes": ComparisonFieldLabel = "General draft notes"
        Case "ObservationStatus", "PromotionRecommendation": ComparisonFieldLabel = AssessmentLabel(key)
        Case "ReportingSeniorName", "ReportingSeniorGrade", "ReportingSeniorDesignator", "ReportingSeniorTitle", "ReportingSeniorUIC": ComparisonFieldLabel = ReportingSeniorLabel(key)
        Case "PromotionStatus": ComparisonFieldLabel = "Promotion status"
        Case "OfficialComments": ComparisonFieldLabel = "Official narrative"
        Case Else
            If Len(PartyNamesLabel(key)) > 0 Then
                ComparisonFieldLabel = PartyNamesLabel(key)
            ElseIf Len(ReportPeopleLabel(key)) > 0 Then
                ComparisonFieldLabel = ReportPeopleLabel(key)
            ElseIf Len(DecisionDetailsLabel(key)) > 0 Then
                ComparisonFieldLabel = DecisionDetailsLabel(key)
            ElseIf Len(ServiceDetailsLabel(key)) > 0 Then
                ComparisonFieldLabel = ServiceDetailsLabel(key)
            ElseIf Len(MemberDetailsLabel(key)) > 0 Then
                ComparisonFieldLabel = MemberDetailsLabel(key)
            ElseIf Len(ReportContextLabel(key)) > 0 Then
                ComparisonFieldLabel = ReportContextLabel(key)
            ElseIf IsTraitKey(key) Then
                ComparisonFieldLabel = Mid$(key, 6) & ". " & CatalogueTraitTitle(Nz(DLookup("FormKind", "Report", "ReportID=" & DLookup("ReportID", "ReportRevision", "RevisionID=" & Forms!frmRevisionCompare!txtNewer)), ""), CLng(Mid$(key, 6)))
                ComparisonFieldLabel = Replace(Replace(ComparisonFieldLabel, vbCr, ""), vbLf, " ")
            Else
                ComparisonFieldLabel = AdministrationLabel(key)
            End If
    End Select
End Function

Public Function LoadComparisonValues() As Boolean
    Dim f As Form, key As String
    Set f = Forms!frmRevisionCompare: key = Nz(f!lstCompareFields, "")
    If key = "" Then Exit Function
    f!txtOldValue = DisplayValue(CLng(f!txtOlder), key)
    f!txtNewValue = DisplayValue(CLng(f!txtNewer), key)
    LoadComparisonValues = True
End Function

Public Function CloseRevisionComparison() As Boolean
    DoCmd.Close acForm, "frmRevisionCompare", acSaveNo
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Forms!frmEditor.SetFocus Else Call OpenScreen("frmLibrary")
    CloseRevisionComparison = True
End Function
