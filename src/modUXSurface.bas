Attribute VB_Name = "modUXSurface"
Option Compare Binary
Option Explicit
Private surfaceBusy As Boolean
Private surfaceSection As Long
Private layoutBusy As Boolean
Private Function Surface() As Access.Form
    Set Surface = Forms!frmEditor!uxFieldSurface.Form
End Function
Public Function UXSurfaceAvailable() As Boolean
    On Error GoTo Missing
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then Exit Function
    Dim f As Access.Form: Set f = Surface()
    UXSurfaceAvailable = True
Missing:
End Function
Public Function UXSurfaceCanonical(ByVal key As String, ByVal value As Variant) As Variant
    If UXFieldFlag(key) Then
        If IsNull(value) Then
            UXSurfaceCanonical = Null
        ElseIf key = "ObservationStatus" Then
            UXSurfaceCanonical = IIf(value, "NOB", "Observed")
        Else
            UXSurfaceCanonical = IIf(value, "Yes", "No")
        End If
    Else
        UXSurfaceCanonical = value
    End If
End Function
Public Function UXSurfacePending(ByVal controlName As String, ByVal fallback As Variant) As Variant
    Dim c As Access.Control, value As Variant, key As String
    UXSurfacePending = fallback
    On Error GoTo Finished
    If Not UXSurfaceAvailable() Then Exit Function
    Set c = Screen.ActiveControl
    If c.Parent.Name <> "frmUXEditorFields" Or c.Name <> controlName Or Left$(c.Tag, 5) <> "SURF:" Then Exit Function
    key = Mid$(c.Tag, 6): value = c.Value
    If c.ControlType = acTextBox Or c.ControlType = acComboBox Then value = c.Text
    UXSurfacePending = UXSurfaceCanonical(key, value)
Finished:
End Function
Public Function UXFlushSurface() As Boolean
    Dim c As Access.Control, key As String, value As Variant
    UXFlushSurface = True
    On Error GoTo Finished
    If Not UXSurfaceAvailable() Then Exit Function
    Set c = Screen.ActiveControl
    If c.Parent.Name <> "frmUXEditorFields" Or Left$(c.Tag, 5) <> "SURF:" Then Exit Function
    key = Mid$(c.Tag, 6): value = UXSurfacePending(c.Name, Forms!frmEditor.Controls(c.Name).Value)
    Forms!frmEditor.Controls(c.Name) = value
Finished:
End Function
Public Function UXSurfaceEdited(ByVal key As String) As Boolean
    If surfaceBusy Then Exit Function
    Dim f As Access.Form, name As String
    Set f = Surface(): name = UXFieldControl(key)
    Forms!frmEditor.Controls(name) = UXSurfaceCanonical(key, f.Controls(name).Value)
    Call UXSurfaceValidate(key)
    If IsTraitKey(key) Then UXSurfaceTrait CLng(Mid$(key, 6))
    UXRefreshEditor
    UXSurfaceEdited = True
End Function
Public Function UXSurfaceTrait(ByVal block As Long) As Boolean
    If Not UXSurfaceAvailable() Then Exit Function
    Call ShowTraitGuidance(block)
    Surface()!uxTraitGuidance = Forms!frmEditor!txtTraitGuidance
    With Surface()!uxTraitGuidance
        .Width = (Surface().Width - 800) \ 2
        .Left = 500 + .Width: .Top = Surface().Controls("cboTrait" & block).Top - 470
        .Height = 2000
    End With
    UXSurfaceTrait = True
End Function
Public Function UXSurfaceValidate(ByVal key As String) As Boolean
    Dim value As Variant, message As String, text As String, digitCount As Long, options As String, item As Variant, allowed As Boolean
    value = UXSurfacePending(UXFieldControl(key), Forms!frmEditor.Controls(UXFieldControl(key)).Value)
    text = Nz(value, "")
    If Len(text) > 0 Then
        If UXFieldLimit(key) > 0 And Len(text) > UXFieldLimit(key) Then message = "Use " & UXFieldLimit(key) & " characters or fewer."
        digitCount = UXFieldDigits(key)
        If digitCount > 0 And (Len(text) <> digitCount Or text Like "*[!0-9]*") Then message = "Enter exactly " & digitCount & " digits."
        If UXFieldDate(key) Then
            If Not IsDate(text) Then message = "Enter a valid date."
        End If
        options = UXFieldOptions(key)
        If Len(options) > 0 Then
            For Each item In Split(options, ";"): If CStr(item) = text Then allowed = True
            Next
            If Not allowed Then message = "Choose a listed value or clear this entry."
        End If
    End If
    If UXSurfaceAvailable() And UXHasControl(Surface(), "err" & key) Then Surface().Controls("err" & key).Caption = message
    UXSurfaceValidate = (Len(message) = 0)
End Function
Public Function UXSurfaceCheckAll() As Boolean
    Dim key As Variant, errors As String, label As String, valid As Boolean
    UXFlushSurface: valid = True
    For Each key In DraftFieldKeys()
        If UXFieldControl(CStr(key)) <> "" And Not UXFieldLegacy(CStr(key)) Then
            If Not UXSurfaceValidate(CStr(key)) Then
                valid = False: label = UXFieldCaption(CStr(key), Nz(Forms!frmEditor!cboKind, "EVAL"))
                If Len(errors) > 0 Then errors = errors & ";"
                errors = errors & CStr(key) & ";" & Replace(label, ";", ",")
            End If
        End If
    Next
    Forms!frmEditor!uxErrorList.RowSourceType = "Value List"
    Forms!frmEditor!uxErrorList.RowSource = errors
    Forms!frmEditor!uxErrorList.Visible = Not valid
    If Not valid Then Forms!frmEditor!uxError.Caption = "Correct these entries before saving. Select an entry to open its section. Pending values are retained."
    UXSurfaceCheckAll = valid
End Function
Public Function UXSurfaceErrorSelected() As Boolean
    Dim key As String
    key = Nz(Forms!frmEditor!uxErrorList, ""): If Len(key) = 0 Then Exit Function
    Call UXSelectSection(UXFieldSection(key))
    Forms!frmEditor!uxFieldSurface.SetFocus
    Surface().Controls(UXFieldControl(key)).SetFocus
    UXSurfaceErrorSelected = True
End Function
Public Function UXSurfaceRefresh() As Boolean
    Dim f As Access.Form, c As Access.Control, key As String, value As Variant, active As String
    If surfaceBusy Or Not UXSurfaceAvailable() Then Exit Function
    On Error GoTo Finished
    Set f = Surface(): surfaceBusy = True
    On Error Resume Next
    If Screen.ActiveControl.Parent.Name = f.Name Then active = Screen.ActiveControl.Name
    On Error GoTo Finished
    For Each c In f.Controls
        If Left$(c.Tag, 5) = "SURF:" And c.Name <> active Then
            key = Mid$(c.Tag, 6): value = Forms!frmEditor.Controls(c.Name).Value
            If UXFieldFlag(key) Then
                If IsNull(value) Or Len(Nz(value, "")) = 0 Then c.Value = Null Else If key = "ObservationStatus" Then c.Value = (value = "NOB") Else c.Value = (value = "Yes")
            Else
                If Not SameFieldValue(c.Value, value) Then c.Value = value
            End If
        End If
    Next
    If UXHasControl(f, "uxNarrativeCount") Then f!uxNarrativeCount.Caption = Len(Nz(Forms!frmEditor!txtOfficialComments, "")) & " characters. Printed fit has not been verified."
    UXSurfaceRefresh = True
Finished:
    surfaceBusy = False
End Function
Private Function UXLayoutBand(ByVal key As String) As String
    Select Case key
        Case "MemberFirstName", "MemberLastName", "MemberMiddleInitial", "MemberSuffix": UXLayoutBand = "name-parts"
        Case "OfficialName", "RankRate", "Designator", "MemberSSN": UXLayoutBand = "member"
        Case "ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW": UXLayoutBand = "component"
        Case "UnitUIC", "ShipStation", "PromotionStatus", "DateReported": UXLayoutBand = "command"
        Case "OccasionPeriodic", "OccasionDetachment", "OccasionFrocking", "OccasionSpecial": UXLayoutBand = "occasion"
        Case "PeriodStart", "PeriodEnd": UXLayoutBand = "period"
        Case "ObservationStatus", "ReportRegular", "ReportConcurrent", "ReportOperationalCommander": UXLayoutBand = "report-kind"
        Case "ReadinessDate", "ReadinessSummary", "ReadinessDetail", "BilletSubcategory": UXLayoutBand = "readiness"
        Case "RetentionYes", "RetentionNo": UXLayoutBand = "retention"
        Case "StatementYes", "StatementNo": UXLayoutBand = "statement"
        Case Else: UXLayoutBand = "details"
    End Select
End Function
Private Function UXLayoutKeys(ByVal section As Long, ByVal kind As String) As Variant
    If section = 0 Then
        UXLayoutKeys = Array("MemberFirstName", "MemberLastName", "MemberMiddleInitial", "MemberSuffix", "OfficialName", "RankRate", "Designator", "MemberSSN", "ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW", "UnitUIC", "ShipStation", "PromotionStatus", "DateReported", "OccasionPeriodic", "OccasionDetachment", "OccasionFrocking", "OccasionSpecial", "PeriodStart", "PeriodEnd", "ObservationStatus", "ReportRegular", "ReportConcurrent", "ReportOperationalCommander", "ReadinessDate", "ReadinessSummary", "ReadinessDetail", "BilletSubcategory")
    Else
        UXLayoutKeys = UXSectionKeys(section, kind)
    End If
End Function
Private Function UXLayoutWidth(ByVal key As String, ByVal available As Long) As Long
    Dim result As Long
    If UXFieldFlag(key) Then
        result = 600 + Len(UXFieldCaption(key, Nz(Forms!frmEditor!cboKind, "EVAL"))) * 80
    Else
        Select Case key
            Case "OfficialName", "ReportingSeniorName", "CounselorName", "ShipStation": result = available \ 2
            Case "MemberMiddleInitial", "MemberSuffix", "RankRate", "UnitUIC", "ReportingSeniorGrade", "ReportingSeniorUIC", "ReportingSeniorState": result = 1600
            Case Else
                If UXFieldDate(key) Then
                    result = 2300
                ElseIf UXFieldLimit(key) > 0 And UXFieldLimit(key) <= 12 Then
                    result = 2200
                Else
                    result = 3500
                End If
        End Select
    End If
    If result > available Then result = available
    UXLayoutWidth = result
End Function
Public Function UXSurfaceSection(ByVal section As Long, ByVal locked As Boolean) As Boolean
    On Error GoTo Failed
    Dim f As Access.Form, c As Access.Control, key As Variant, kind As String, row As Long, slot As Long, top As Long, x As Long, multi As Boolean, applicable As Boolean, tabOrder As Long, columns As Long, cellWidth As Long, usableWidth As Long, band As String, priorBand As String, rowHeight As Long, fieldWidth As Long, flag As Boolean
    Set f = Surface(): surfaceSection = section: kind = Nz(Forms!frmEditor!cboKind, "EVAL")
    'Access keeps a subform's last focused field even when focus leaves the
    'parent. Move to a permanent anchor before hiding any section controls.
    usableWidth = Forms!frmEditor!uxFieldSurface.Width - 800
    columns = IIf(usableWidth >= 14000, 3, 2)
    If section = 4 Then columns = 2
    cellWidth = (usableWidth - (columns - 1) * 300) \ columns
    f!uxFocusAnchor.SetFocus
    For Each c In f.Controls
        If c.Name <> "uxFocusAnchor" Then
            c.Visible = False: c.Left = 200: c.Width = 200
        End If
    Next
    top = 200: x = 200
    For Each key In UXLayoutKeys(section, kind)
        If Not UXFieldLegacy(key) And key <> "FormKind" Then
            multi = UXFieldMultiline(key): applicable = UXFieldApplies(key, kind): flag = UXFieldFlag(key)
            band = UXLayoutBand(CStr(key))
            fieldWidth = UXLayoutWidth(CStr(key), usableWidth)
            If section = 4 Then fieldWidth = cellWidth: band = CStr(key)
            If multi Then fieldWidth = usableWidth: band = CStr(key)
            If x > 200 And (band <> priorBand Or x + fieldWidth > usableWidth + 200) Then
                top = top + rowHeight: x = 200: rowHeight = 0
            End If
            Set c = f.Controls("label" & key): c.Caption = UXFieldCaption(key, kind): c.Left = x: c.Top = top: c.Width = fieldWidth: c.Height = 420: c.Visible = True
            Set c = f.Controls(UXFieldControl(key)): c.Left = x: c.Top = top + 450: c.Width = fieldWidth: c.Height = 420: c.Visible = True: c.Locked = locked Or Not applicable
            c.TabIndex = tabOrder: tabOrder = tabOrder + 1
            If flag Then
                c.Top = top: c.Width = 320: c.Height = 320
                With f.Controls("label" & key)
                    .Left = x + 420: .Width = fieldWidth - 420
                End With
            End If
            If multi Then c.Height = 2600
            Set c = f.Controls("err" & key): c.Left = x: c.Top = top + IIf(flag, 350, 890): c.Width = fieldWidth: c.Height = 250: c.Visible = True
            If Not applicable Then c.Caption = "Not used on this report type; stored value retained." Else UXSurfaceValidate key
            If multi Then c.Top = top + 3100
            If multi Then
                rowHeight = 3500
            ElseIf flag Then
                If rowHeight < 650 Then rowHeight = 650
            Else
                If rowHeight < 1200 Then rowHeight = 1200
            End If
            x = x + fieldWidth + 250: priorBand = band
        End If
    Next
    top = top + rowHeight + 200
    f.Width = Forms!frmEditor!uxFieldSurface.Width - 300
    If section = 4 Then
        f!uxTraitGuidance.Visible = True: f!uxTraitGuidance.TabIndex = tabOrder
        f!uxTraitGuidance.Left = 200 + cellWidth + 300
        f!uxTraitGuidance.Width = cellWidth
        top = top + 2000: UXSurfaceTrait 33
    End If
    If section = 5 Then
        f!uxNarrativeCount.Top = top: f!uxNarrativeCount.Width = usableWidth: f!uxNarrativeCount.Visible = True
        f!cmdSnippets.Width = 3000: f!cmdSnippets.Top = top + 450: f!cmdSnippets.Visible = True: f!cmdSnippets.Enabled = Not locked
        top = top + 1100
    End If
    f.Section(acDetail).Height = top + 200
    Forms!frmEditor!uxFieldSurface.Visible = (section <= 7)
    UXSurfaceRefresh
    UXSurfaceSection = True
    Exit Function
Failed:
    Forms!frmEditor!uxError.Caption = "Section could not open: " & Err.Number & " " & Err.Description
End Function

Public Function UXResizeEditor(Optional ByVal viewportWidth As Long = 0, Optional ByVal viewportHeight As Long = 0) As Boolean
    Dim f As Access.Form, width As Long, height As Long, contentLeft As Long, contentWidth As Long, footer As Long, i As Long
    If layoutBusy Or Not UXSurfaceAvailable() Then Exit Function
    On Error GoTo Failed
    layoutBusy = True: Set f = Forms!frmEditor
    UXFlushSurface
    width = f.InsideWidth: height = f.InsideHeight
    If viewportWidth > 0 Then width = viewportWidth
    If viewportHeight > 0 Then height = viewportHeight
    If width < 12000 Or height < 7500 Then GoTo Finished
    contentLeft = 5900: contentWidth = width - contentLeft - 350: footer = height - 1300
    For i = 0 To 9
        With f.Controls("uxSection" & i)
            .Left = 2800: .Top = 3250 + i * 430: .Width = 2800: .Height = 400
        End With
    Next
    f!uxRail.Height = height - 100
    f!uxContext.Width = width - 3200
    f!uxITA.Width = width - 3200
    f!uxSGA.Width = width - 8000
    f!uxGroup.Left = width - 4350: f!uxGroup.Width = 4000: f!uxGroup.Height = 450
    f!uxLocked.Left = 2850: f!uxLocked.Top = 2350: f!uxLocked.Width = width - 3200: f!uxLocked.Height = 400
    f!uxSection.Left = contentLeft: f!uxSection.Top = 2800: f!uxSection.Width = contentWidth
    f!uxError.Left = contentLeft: f!uxError.Top = 3250: f!uxError.Width = contentWidth
    f!uxErrorList.Left = contentLeft: f!uxErrorList.Top = 3550: f!uxErrorList.Width = contentWidth
    f!uxFieldSurface.Left = contentLeft
    f!uxFieldSurface.Top = IIf(surfaceSection = 1, 4100, 4000)
    f!uxFieldSurface.Width = contentWidth
    f!uxFieldSurface.Height = footer - f!uxFieldSurface.Top - 550
    f!lblSaved.Left = contentLeft: f!lblSaved.Top = footer - 400: f!lblSaved.Width = contentWidth
    f!cmdClose.Left = contentLeft: f!cmdClose.Top = footer: f!cmdClose.Height = 500
    f!cmdSave.Left = width - 2850: f!cmdSave.Top = footer: f!cmdSave.Height = 500
    f!uxDirty.Left = contentLeft + 2600: f!uxDirty.Top = footer: f!uxDirty.Width = contentWidth - 5500
    f!lblRecovery.Left = contentLeft: f!lblRecovery.Top = footer + 650: f!lblRecovery.Width = contentWidth
    f!uxSeniorRecord.Left = contentLeft: f!uxSeniorRecord.Top = 3500
    f!uxUseSenior.Left = contentLeft + 5000: f!uxUseSenior.Top = 3500
    UXSurfaceSection surfaceSection, Not f!cmdSave.Enabled
    UXResizeEditor = True
Finished:
    layoutBusy = False
    Exit Function
Failed:
    f!uxError.Caption = "Layout could not resize: " & Err.Description
    Resume Finished
End Function
