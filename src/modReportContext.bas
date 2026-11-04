Attribute VB_Name = "modReportContext"
Option Compare Binary
Option Explicit
Public Function ReportContextKeys() As Variant
    ReportContextKeys = Array("OccasionPeriodic", "OccasionDetachment", "OccasionFrocking", "OccasionSpecial", "ReportRegular", "ReportConcurrent", "ReportOperationalCommander", "PrimaryDuty", "DutyDescription")
End Function
Public Function ReportContextPanelKeys(ByVal panel As String) As Variant
    Select Case panel
        Case "Occasions": ReportContextPanelKeys = Array("OccasionPeriodic", "OccasionDetachment", "OccasionFrocking", "OccasionSpecial")
        Case "Categories": ReportContextPanelKeys = Array("ReportRegular", "ReportConcurrent", "ReportOperationalCommander")
        Case "Duties": ReportContextPanelKeys = Array("PrimaryDuty", "DutyDescription")
        Case Else: Err.Raise vbObjectError + 2500, "modReportContext.ReportContextPanelKeys", "Unknown report context panel."
    End Select
End Function
Public Function ReportContextLabel(ByVal key As String) As String
    Select Case key
        Case "OccasionPeriodic": ReportContextLabel = "Periodic occasion"
        Case "OccasionDetachment": ReportContextLabel = "Detachment of individual occasion"
        Case "OccasionFrocking": ReportContextLabel = "Frocking occasion"
        Case "OccasionSpecial": ReportContextLabel = "Special occasion"
        Case "ReportRegular": ReportContextLabel = "Regular report"
        Case "ReportConcurrent": ReportContextLabel = "Concurrent report"
        Case "ReportOperationalCommander": ReportContextLabel = "Operational commander report"
        Case "PrimaryDuty": ReportContextLabel = "Primary duty"
        Case "DutyDescription": ReportContextLabel = "Duty description"
    End Select
End Function
Public Function ValidateReportContext(ByVal values As Variant) As Variant
    Dim i As Long, choice As Variant, allowed As Boolean
    If Not IsArray(values) Then Err.Raise vbObjectError + 2501, "modReportContext.ValidateReportContext", "Report context requires nine text or null entries."
    If LBound(values) <> 0 Or UBound(values) <> 8 Then Err.Raise vbObjectError + 2501, "modReportContext.ValidateReportContext", "Report context requires nine text or null entries."
    For i = 0 To 8
        If Not IsNull(values(i)) Then
            If VarType(values(i)) <> vbString Then Err.Raise vbObjectError + 2502, "modReportContext.ValidateReportContext", "Report context entries must be text or null."
            If i < 7 Then
                allowed = False
                For Each choice In Array("", "Yes", "No")
                    If values(i) = choice Then allowed = True
                Next
                If Not allowed Then Err.Raise vbObjectError + 2503, "modReportContext.ValidateReportContext", "Choose Yes, No or blank for each report-context flag."
            ElseIf i = 7 Then
                If Len(values(i)) > UXFieldLimit("PrimaryDuty") Or InStr(values(i), vbCr) > 0 Or InStr(values(i), vbLf) > 0 Or InStr(values(i), vbTab) > 0 Then Err.Raise vbObjectError + 2504, "modReportContext.ValidateReportContext", "Primary duty exceeds its field limit or contains line breaks/tabs."
            End If
        End If
    Next
    ValidateReportContext = values
End Function
Public Function ReportContextFromEditor() As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null, Null, Null): keys = ReportContextKeys()
    For i = 0 To 8: values(i) = Forms!frmEditor.Controls("txt" & keys(i)).Value: Next
    ReportContextFromEditor = values
End Function
Public Function ReportContextFromRecordset(ByVal r As DAO.Recordset) As Variant
    Dim values As Variant, keys As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null, Null, Null): keys = ReportContextKeys()
    For i = 0 To 8: values(i) = r.Fields(keys(i)).Value: Next
    ReportContextFromRecordset = values
End Function
Public Function ReportContextFromRevision(ByVal revisionID As Long) As Variant
    Dim values As Variant, keys As Variant, r As DAO.Recordset, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null, Null, Null): keys = ReportContextKeys()
    For i = 0 To 8
        Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='" & keys(i) & "'", dbOpenSnapshot)
        If r.EOF Then
            If RevisionSnapshotSchema(revisionID) >= 14 Then r.Close: Err.Raise vbObjectError + 2505, "modReportContext.ReportContextFromRevision", "Report context snapshot is incomplete."
        ElseIf Not r!IsNullValue Then
            If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 2506, "modReportContext.ReportContextFromRevision", "Report context snapshot has inconsistent null flag."
            values(i) = r!ValueText
        End If
        r.Close
    Next
    ReportContextFromRevision = ValidateReportContext(values)
End Function
Public Function OpenReportContext(ByVal panel As String) As Boolean
    Dim keys As Variant, key As Variant, screen As String
    keys = ReportContextPanelKeys(panel): screen = "frmContext" & panel
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 2507, "modReportContext.OpenReportContext", "Open a draft first."
    DoCmd.OpenForm screen
    Call UXAssertView(screen)
    For Each key In keys
        Forms(screen).Controls("txt" & key) = Forms!frmEditor.Controls("txt" & key).Value
    Next
    Forms(screen)!cmdApply.Enabled = Forms!frmEditor!cmdSave.Enabled
    Forms(screen).SetFocus: OpenReportContext = True
End Function
Public Function ApplyReportContext(ByVal panel As String) As Boolean
    Dim key As Variant, screen As String, keys As Variant
    keys = ReportContextPanelKeys(panel): screen = "frmContext" & panel
    If Not Forms!frmEditor!cmdSave.Enabled Then Err.Raise vbObjectError + 2508, "modReportContext.ApplyReportContext", "Recover the archived draft before editing."
    Call UXAssertView(screen)
    For Each key In keys
        Forms!frmEditor.Controls("txt" & key) = Forms(screen).Controls("txt" & key).Value
    Next
    DoCmd.Close acForm, screen, acSaveNo
    Call CheckpointEditor()
    Forms!frmEditor.SetFocus: ApplyReportContext = True
End Function
Public Function CancelReportContext(ByVal panel As String) As Boolean
    Dim keys As Variant
    keys = ReportContextPanelKeys(panel)
    DoCmd.Close acForm, "frmContext" & panel, acSaveNo
    Forms!frmEditor.SetFocus: CancelReportContext = True
End Function
