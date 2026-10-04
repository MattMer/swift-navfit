Attribute VB_Name = "modUXSurface"
Option Compare Binary
Option Explicit
Private surfaceBusy As Boolean
Private surfaceSection As Long
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
        .Left = 4100: .Top = Surface().Controls("cboTrait" & block).Top - 470
        .Width = 3600: .Height = 2000
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
Public Function UXSurfaceSection(ByVal section As Long, ByVal locked As Boolean) As Boolean
    On Error GoTo Failed
    Dim f As Access.Form, c As Access.Control, key As Variant, kind As String, row As Long, slot As Long, top As Long, x As Long, multi As Boolean, applicable As Boolean, tabOrder As Long
    Set f = Surface(): surfaceSection = section: kind = Nz(Forms!frmEditor!cboKind, "EVAL")
    'Access keeps a subform's last focused field even when focus leaves the
    'parent. Move to a permanent anchor before hiding any section controls.
    f!uxFocusAnchor.SetFocus
    For Each c In f.Controls
        If c.Name <> "uxFocusAnchor" Then c.Visible = False
    Next
    For Each key In UXSectionKeys(section, kind)
        If Not UXFieldLegacy(key) And key <> "FormKind" Then
            multi = UXFieldMultiline(key): applicable = UXFieldApplies(key, kind)
            If multi And slot Mod 2 = 1 Then slot = slot + 1
            top = 200 + (slot \ 2) * 1250: x = 200 + (slot Mod 2) * 3900
            Set c = f.Controls("label" & key): c.Caption = UXFieldCaption(key, kind): c.Left = x: c.Top = top: c.Width = 3600: c.Height = 450: c.Visible = True
            Set c = f.Controls(UXFieldControl(key)): c.Left = x: c.Top = top + 470: c.Width = 3600: c.Height = 450: c.Visible = True: c.Locked = locked Or Not applicable
            c.TabIndex = tabOrder: tabOrder = tabOrder + 1
            If UXFieldFlag(key) Then c.Width = 320: c.Height = 320
            If multi Then c.Width = 7500: c.Height = 2600: f.Controls("label" & key).Width = 7500
            Set c = f.Controls("err" & key): c.Left = x: c.Top = top + 950: c.Width = 3600: c.Height = 250: c.Visible = True
            If Not applicable Then c.Caption = "Not used on this report type; stored value retained." Else UXSurfaceValidate key
            If multi Then
                c.Top = top + 3120: c.Width = 7500
                slot = slot + 6
            Else
                slot = slot + IIf(section = 4, 2, 1)
            End If
        End If
    Next
    top = 200 + ((slot + 1) \ 2) * 1250
    If section = 4 Then
        f!uxTraitGuidance.Visible = True: f!uxTraitGuidance.TabIndex = tabOrder
        top = top + 2000: UXSurfaceTrait 33
    End If
    If section = 5 Then
        f!uxNarrativeCount.Top = top: f!uxNarrativeCount.Visible = True
        f!cmdSnippets.Top = top + 450: f!cmdSnippets.Visible = True: f!cmdSnippets.Enabled = Not locked
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
