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
    Dim f As Access.Form, name As String, hadError As Boolean
    Set f = Surface(): name = UXFieldControl(key)
    Forms!frmEditor.Controls(name) = UXSurfaceCanonical(key, f.Controls(name).Value)
    hadError = Len(f.Controls("err" & key).Caption) > 0
    Call UXSurfaceValidate(key)
    If hadError <> (Len(f.Controls("err" & key).Caption) > 0) Then
        UXResizeEditor
        f.Controls(name).SetFocus
    End If
    If IsTraitKey(key) Then UXSurfaceTrait CLng(Mid$(key, 6))
    UXRefreshEditor
    UXSurfaceEdited = True
End Function
Public Function UXSurfaceTrait(ByVal block As Long) As Boolean
    If Not UXSurfaceAvailable() Then Exit Function
    Call ShowTraitGuidance(block)
    Surface()!uxTraitGuidance = Forms!frmEditor!txtTraitGuidance
    With Surface()!uxTraitGuidance
        If Surface().Width < 7000 Then
            .Left = 200: .Width = Surface().Width - 500
            .Top = Surface().Controls("cboTrait39").Top + 700
        Else
            .Left = Surface().Controls("cboTrait" & block).Left + 1450
            .Width = Surface().Width - .Left - 300: .Top = 200
        End If
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
    If UXSurfaceAvailable() And UXHasControl(Surface(), "err" & key) Then
        With Surface().Controls("err" & key)
            .Caption = message: .Visible = (Len(message) > 0 And Surface().Controls(UXFieldControl(key)).Visible)
        End With
    End If
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
    UXResizeEditor
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
Public Function UXLayoutBand(ByVal key As String) As String
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
Public Function UXLayoutKeys(ByVal section As Long, ByVal kind As String) As Variant
    If section = 0 Then
        UXLayoutKeys = Array("MemberFirstName", "MemberLastName", "MemberMiddleInitial", "MemberSuffix", "OfficialName", "RankRate", "Designator", "MemberSSN", "ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW", "UnitUIC", "ShipStation", "PromotionStatus", "DateReported", "OccasionPeriodic", "OccasionDetachment", "OccasionFrocking", "OccasionSpecial", "PeriodStart", "PeriodEnd", "ObservationStatus", "ReportRegular", "ReportConcurrent", "ReportOperationalCommander", "ReadinessDate", "ReadinessSummary", "ReadinessDetail", "BilletSubcategory")
    Else
        UXLayoutKeys = UXSectionKeys(section, kind)
    End If
End Function
Public Function UXLayoutWidth(ByVal key As String, ByVal available As Long, Optional ByVal kind As String = "") As Long
    Dim result As Long
    If Len(kind) = 0 Then kind = Nz(Forms!frmEditor!cboKind, "EVAL")
    If UXFieldFlag(key) Then
        result = 550 + Len(UXFieldCaption(key, kind)) * 100
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
    If Not UXFieldMultiline(key) And Not IsTraitKey(key) Then
        If result < Len(UXFieldCaption(key, kind)) * 100 Then result = Len(UXFieldCaption(key, kind)) * 100
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
    cellWidth = (usableWidth - 300) \ 2
    If section = 4 And usableWidth < 6500 Then cellWidth = usableWidth
    If cellWidth > 6500 Then cellWidth = 6500
    f!uxFocusAnchor.SetFocus
    For Each c In f.Controls
        If c.Name <> "uxFocusAnchor" Then
            c.Visible = False: c.Left = 200: c.Top = 200: c.Width = 200
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
            Set c = f.Controls("label" & key): c.Caption = UXFieldCaption(key, kind): c.Left = x: c.Top = top: c.Width = fieldWidth: c.Height = 300: c.Visible = True
            Set c = f.Controls(UXFieldControl(key)): c.Left = x: c.Top = top + 320: c.Width = fieldWidth: c.Height = 360: c.Visible = True: c.Locked = locked Or Not applicable
            c.TabIndex = tabOrder: tabOrder = tabOrder + 1
            If flag Then
                c.Top = top: c.Width = 280: c.Height = 280
                With f.Controls("label" & key)
                    .Left = x + 360: .Width = fieldWidth - 360
                End With
            End If
            If multi Then c.Height = 2600
            Set c = f.Controls("err" & key): c.Left = x: c.Top = top + IIf(flag, 300, 700): c.Width = fieldWidth: c.Height = 250: c.Visible = True
            If Not applicable Then c.Caption = "Not used on this report type; stored value retained." Else UXSurfaceValidate key
            If multi Then c.Top = top + 2960
            If multi Then
                rowHeight = 3240
            ElseIf flag Then
                If rowHeight < 380 Then rowHeight = 380
            Else
                If rowHeight < 760 Then rowHeight = 760
            End If
            If section = 4 Then
                With f.Controls("label" & key)
                    .Width = cellWidth - 1450: .Height = 360
                End With
                With f.Controls(UXFieldControl(key))
                    .Left = x + cellWidth - 1300: .Top = top: .Width = 1000: .Height = 360
                End With
                f.Controls("err" & key).Top = top + 390
                rowHeight = 600
            End If
            If Len(f.Controls("err" & key).Caption) > 0 Then
                If rowHeight < f.Controls("err" & key).Top - top + 300 Then rowHeight = f.Controls("err" & key).Top - top + 300
            End If
            x = x + fieldWidth + 160: priorBand = band
        End If
    Next
    top = top + rowHeight + 200
    f.Width = Forms!frmEditor!uxFieldSurface.Width - 300
    If section = 4 Then
        f!uxTraitGuidance.Visible = True: f!uxTraitGuidance.TabIndex = tabOrder
        f!uxTraitGuidance.Left = 200 + cellWidth + 300
        f!uxTraitGuidance.Width = cellWidth
        UXSurfaceTrait 33
        If usableWidth < 6500 Then
            f!uxTraitGuidance.Left = 200: f!uxTraitGuidance.Top = top
            f!uxTraitGuidance.Width = usableWidth: top = top + 2200
        End If
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
    Dim f As Access.Form, width As Long, height As Long, contentLeft As Long, contentWidth As Long, footer As Long, i As Long, actions As Variant, active As Access.Control, caret As Long, selection As Long, compact As Boolean, surfaceTop As Long
    If layoutBusy Or Not UXSurfaceAvailable() Then Exit Function
    On Error GoTo Failed
    layoutBusy = True: Set f = Forms!frmEditor
    On Error Resume Next
    Set active = Screen.ActiveControl
    caret = active.SelStart: selection = active.SelLength
    On Error GoTo Failed
    UXFlushSurface
    width = f.InsideWidth: height = f.InsideHeight
    If viewportWidth > 0 Then width = viewportWidth
    If viewportHeight > 0 Then height = viewportHeight
    If width < 4000 Or height < 2000 Then GoTo Finished
    compact = (width < 12000 Or height < 7500)
    UXNavigationLayout f, compact, width
    f!uxCompactSection.Visible = compact: f!uxCompactSection = surfaceSection
    If compact Then
        UXCompactEditorGeometry f, width, height
        GoTo Render
    End If
    f.ScrollBars = 0
    f!uxContext.Left = 2850: f!uxContext.Top = 300
    f!uxReport.Left = 2850: f!uxReport.Top = 850: f!uxReport.Width = width - 5900: f!uxReport.Height = 650
    f!cboKind.Left = width - 2850: f!cboKind.Top = 1000: f!cboKind.Width = 2500
    f!uxTypeCaption.Left = width - 4650: f!uxTypeCaption.Top = 650: f!uxTypeCaption.Width = 4300
    f!uxITA.Left = 2850: f!uxITA.Top = 1800: f!uxITA.Height = 400
    f!uxSGA.Left = 2850: f!uxSGA.Top = 2200: f!uxSGA.Height = 400
    f!uxGroup.Top = 2200
    f!uxAlias.Visible = True: f!uxAlias.Left = 2850: f!uxAlias.Top = 1500: f!uxAlias.Width = width - 3200: f!uxAlias.Height = 280
    contentLeft = 5900: contentWidth = width - contentLeft - 350: footer = height - 1300
    For i = 0 To 9
        With f.Controls("uxSection" & i)
            .Visible = True: .Left = 2800: .Top = 3250 + i * 430: .Width = 2800: .Height = 400
        End With
    Next
    f!uxRail.Height = height - 100
    f!uxContext.Width = width - 3200
    f!uxITA.Width = width - 3200
    f!uxSGA.Width = width - 8000
    f!uxGroup.Left = width - 4350: f!uxGroup.Width = 4000: f!uxGroup.Height = 450
    f!uxLocked.Left = contentLeft: f!uxLocked.Top = 2650: f!uxLocked.Width = contentWidth: f!uxLocked.Height = 400
    f!uxSection.Left = contentLeft: f!uxSection.Top = 2800: f!uxSection.Width = contentWidth
    If Len(f!uxLocked.Caption) > 0 Then f!uxSection.Top = 3100
    f!uxError.Left = contentLeft: f!uxError.Top = 3250: f!uxError.Width = contentWidth
    f!uxErrorList.Left = contentLeft: f!uxErrorList.Top = 3550: f!uxErrorList.Width = contentWidth
    f!uxFieldSurface.Left = contentLeft
    f!uxFieldSurface.Top = IIf(surfaceSection = 1, 4100, IIf(f!uxErrorList.Visible, 4100, 3350))
    If Len(f!uxLocked.Caption) > 0 Then f!uxFieldSurface.Top = f!uxFieldSurface.Top + 400
    f!uxFieldSurface.Width = contentWidth
    f!uxFieldSurface.Height = footer - f!uxFieldSurface.Top - 550
    f!lblSaved.Left = contentLeft: f!lblSaved.Top = footer - 400: f!lblSaved.Width = contentWidth
    f!cmdClose.Left = contentLeft: f!cmdClose.Top = footer: f!cmdClose.Height = 500
    f!cmdSave.Left = width - 2850: f!cmdSave.Top = footer: f!cmdSave.Height = 500
    f!cmdSave.Caption = "Save draft": f!cmdSave.ForeColor = 5916729
    f!uxDirty.Left = contentLeft + 2600: f!uxDirty.Top = footer: f!uxDirty.Width = contentWidth - 5500
    f!lblRecovery.Left = contentLeft: f!lblRecovery.Top = footer + 650: f!lblRecovery.Width = contentWidth
    If surfaceSection = 8 Then
        actions = Array("cmdRoute", "cmdComments", "cmdChecklist", "cmdCalculations")
        For i = 0 To UBound(actions)
            With f.Controls(actions(i))
                .Left = contentLeft + i * (contentWidth \ 4): .Top = 3350 + IIf(Len(f!uxLocked.Caption) > 0, 400, 0): .Width = contentWidth \ 4 - 200: .Height = 420
            End With
        Next
    ElseIf surfaceSection = 9 Then
        f!cmdCompare.Left = contentLeft: f!cmdCompare.Top = 3350: f!cmdCompare.Width = 2300: f!cmdCompare.Height = 420
        f!cmdRestore.Left = contentLeft + 2500: f!cmdRestore.Top = 3350: f!cmdRestore.Width = 2300: f!cmdRestore.Height = 420
        f!lstRevisions.Left = contentLeft: f!lstRevisions.Top = 3950: f!lstRevisions.Width = contentWidth: f!lstRevisions.Height = footer - 4500
        If Len(f!uxLocked.Caption) > 0 Then
            f!cmdCompare.Top = f!cmdCompare.Top + 400: f!cmdRestore.Top = f!cmdRestore.Top + 400
            f!lstRevisions.Top = f!lstRevisions.Top + 400: f!lstRevisions.Height = f!lstRevisions.Height - 400
        End If
    End If
    f!uxSeniorRecord.Left = contentLeft: f!uxSeniorRecord.Top = 3500: f!uxSeniorRecord.Width = contentWidth - 3300
    f!uxUseSenior.Left = width - 3400: f!uxUseSenior.Top = 3500: f!uxUseSenior.Width = 3050
    f.Width = width - 100
    f.Section(acDetail).Height = height - 100
Render:
    f!cmdSave.Caption = "Save draft": f!cmdSave.ForeColor = 5916729
    If Not UXSurfaceSection(surfaceSection, Not f!cmdSave.Enabled) Then GoTo Finished
    'Reflow hides controls temporarily. Restore the original field and caret,
    'including raw text still being typed, without committing a revision.
    On Error Resume Next
    If Not active Is Nothing Then
        If active.Visible And active.Enabled Then
            If active.Parent.Name = "frmUXEditorFields" Then f!uxFieldSurface.SetFocus
            active.SetFocus
            active.SelStart = caret: active.SelLength = selection
            If active.Parent.Name = "frmUXEditorFields" And surfaceSection = 4 Then UXSurfaceTrait CLng(Mid$(active.Name, 9))
        End If
    End If
    On Error GoTo Failed
    UXResizeEditor = True
Finished:
    layoutBusy = False
    Exit Function
Failed:
    f!uxError.Caption = "Layout could not resize: " & Err.Description
    Resume Finished
End Function

Private Sub UXEditorPosition(ByVal f As Access.Form, ByVal name As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    With f.Controls(name)
        .Left = x: .Top = y: .Width = width: .Height = height
    End With
End Sub
Private Sub UXCompactEditorGeometry(ByVal f As Access.Form, ByVal width As Long, ByVal height As Long)
    Dim i As Long, available As Long, footer As Long, top As Long, effectiveHeight As Long, actions As Variant
    available = width - 600: effectiveHeight = height
    If effectiveHeight < 6000 Then effectiveHeight = 6000
    f.ScrollBars = IIf(height < 6000, 2, 0)
    f!uxCompactNav.SetFocus
    For i = 0 To 9: f.Controls("uxSection" & i).Visible = False: Next
    UXEditorPosition f, "uxContext", 2950, 150, width - 3300, 500
    UXEditorPosition f, "uxReport", 250, 700, available - 2500, 650
    UXEditorPosition f, "uxTypeCaption", width - 2600, 650, 2250, 300
    UXEditorPosition f, "cboKind", width - 2600, 1000, 2250, 400
    UXEditorPosition f, "uxITA", 250, 1400, available, 400
    UXEditorPosition f, "uxSGA", 250, 1820, available - 2900, 650
    UXEditorPosition f, "uxGroup", width - 3100, 1820, 2750, 450
    UXEditorPosition f, "uxLocked", 250, 2470, available, 400
    UXEditorPosition f, "uxAlias", 250, 2880, available, 350
    UXEditorPosition f, "uxCompactSection", 250, 3250, available, 420
    UXEditorPosition f, "uxSection", 250, 3710, available, 420
    'Avoid a blank alias/lock row while retaining real status text.
    top = 2470
    If Len(f!uxLocked.Caption) > 0 Then f!uxLocked.Top = top: top = top + 420
    f!uxAlias.Top = top
    If Len(f!uxAlias.Caption) > 0 Then top = top + 350
    f!uxCompactSection.Top = top: top = top + 450
    f!uxSection.Top = top: top = top + 450
    UXEditorPosition f, "uxError", 250, top, available, 300
    If Len(f!uxError.Caption) > 0 Then top = top + 330
    UXEditorPosition f, "uxErrorList", 250, top, available, 430
    If f!uxErrorList.Visible Then top = top + 460
    If surfaceSection = 1 Then
        UXEditorPosition f, "uxSeniorRecord", 250, top, available - 2300, 420
        UXEditorPosition f, "uxUseSenior", width - 2450, top, 2100, 420
        top = top + 480
    End If
    'Grow the scrollable document when status/error rows leave too little space.
    If effectiveHeight < top + 2400 Then effectiveHeight = top + 2400: f.ScrollBars = 2
    footer = effectiveHeight - 1150
    UXEditorPosition f, "uxFieldSurface", 250, top, available, footer - top - 450
    UXEditorPosition f, "lblSaved", 250, footer - 400, available, 300
    UXEditorPosition f, "cmdClose", 250, footer, 2300, 500
    UXEditorPosition f, "cmdSave", width - 2650, footer, 2300, 500
    If available >= 5200 Then
        UXEditorPosition f, "uxDirty", 2700, footer, available - 5000, 500
    Else
        UXEditorPosition f, "cmdClose", 250, footer, available \ 2 - 100, 500
        UXEditorPosition f, "cmdSave", 350 + available \ 2, footer, available \ 2 - 100, 500
        UXEditorPosition f, "uxDirty", 250, footer + 520, available, 300
    End If
    UXEditorPosition f, "lblRecovery", 250, footer + 550, available, 500
    If surfaceSection = 8 Or surfaceSection = 9 Then
        If surfaceSection = 8 Then actions = Array("cmdRoute", "cmdComments", "cmdChecklist", "cmdCalculations") Else actions = Array("cmdCompare", "cmdRestore")
        For i = 0 To UBound(actions)
            UXEditorPosition f, CStr(actions(i)), 250 + i * (available \ (UBound(actions) + 1)), top, available \ (UBound(actions) + 1) - 150, 500
        Next
        If surfaceSection = 9 Then UXEditorPosition f, "lstRevisions", 250, top + 600, available, footer - top - 1050
    End If
    f.Width = width - 100: f.Section(acDetail).Height = effectiveHeight - 100
End Sub
