Attribute VB_Name = "modUXViewport"
Option Explicit
Private layouts As Object
Private viewportBusy As Boolean
Private viewportError As String
Private layoutPasses As Long
' Widest content area a workspace stretches to (twips right of the rail); wider
' windows leave a right margin instead of over-long rows.
Private Const MAX_CONTENT_WIDTH As Long = 16500
' Access rejects form and control positions past 22 inches; a maximized window on
' a 2560-pixel or scaled 4K screen is wider than that.
Public Const UX_MAX_FORM_WIDTH As Long = 31680
Public Function UXViewportPasses() As Long
    UXViewportPasses = layoutPasses
End Function
Public Function UXViewportFailure() As String
    UXViewportFailure = viewportError
End Function
Public Function UXCompactNavigate(ByVal formName As String) As Boolean
    Dim target As String
    target = Nz(Forms(formName)!uxCompactNav, "")
    If Len(target) > 0 Then UXCompactNavigate = UXNavigate(target)
End Function
Public Function UXChooseSectionCompact() As Boolean
    UXChooseSectionCompact = UXSelectSection(Nz(Forms!frmEditor!uxCompactSection, 0))
End Function
Public Sub UXNavigationLayout(ByVal f As Access.Form, ByVal compact As Boolean, ByVal width As Long)
    Dim i As Long, active As Access.Control, origin As Access.Form
    On Error Resume Next
    Set origin = Screen.ActiveForm
    Set active = f.ActiveControl
    On Error GoTo 0
    If compact Then UXVisualSet f!uxCompactNav, "Visible", True
    For i = 0 To 5
        If Not compact Then UXVisualSet f.Controls("uxNav" & i), "Visible", True
    Next
    If Not active Is Nothing Then
    If compact And active.Parent.Name = f.Name Then
        If active.Name Like "uxNav#" Or active.Name Like "uxSection#" Then
            f.SetFocus: f!uxCompactNav.SetFocus
        End If
    ElseIf Not compact And active.Parent.Name = f.Name Then
        If active.Name = "uxCompactNav" Or active.Name = "uxCompactSection" Then
            f.SetFocus: f!uxNav0.SetFocus
        End If
    End If
    End If
    If Not origin Is Nothing Then
        If origin.Name <> f.Name Then origin.SetFocus
    End If
    On Error GoTo 0
    UXVisualSet f!uxCompactNav, "Visible", compact
    UXVisualSet f!uxRail, "Visible", Not compact: UXVisualSet f!uxBrand, "Visible", Not compact: UXVisualSet f!uxNavCaption, "Visible", Not compact
    For i = 0 To 5: UXVisualSet f.Controls("uxNav" & i), "Visible", Not compact: Next
    If compact Then
        Select Case f.Name
            Case "frmEditor", "frmLibrary": f!uxCompactNav = "frmLibrary"
            Case "frmMajorWorkbench", "frmUXBulk", "frmUXComparison", "frmUXGroupRSCA", "frmUXGroups": f!uxCompactNav = "frmUXGroups"
            Case "frmUXSeniors", "frmUXSeniorEditor", "frmUXResources", "frmUXRouteTemplates": f!uxCompactNav = "frmUXResources"
            Case "frmMyWork": f!uxCompactNav = "frmMyWork"
            Case "frmUXSettings": f!uxCompactNav = "frmUXSettings"
            Case Else: f!uxCompactNav = "frmHome"
        End Select
        f!uxCompactNav.Left = 250: f!uxCompactNav.Top = 150
        f!uxCompactNav.Width = 2500: f!uxCompactNav.Height = 400
    End If
End Sub
Public Sub UXRestoreViewport(ByVal f As Access.Form)
    Dim state As Object, geometry As Object, c As Access.Control, v As Variant, old As Access.Form
    If layouts Is Nothing Then Exit Sub
    If Not layouts.Exists(f.Name) Then Exit Sub
    Set state = layouts(f.Name): Set old = state("Form")
    If old Is f Then
        Set geometry = state("Geometry")
        For Each c In f.Controls
            If geometry.Exists(c.Name) Then
                v = geometry(c.Name): c.Left = v(0): c.Top = v(1): c.Width = v(2): c.Height = v(3)
            End If
        Next
        f.Section(acDetail).Height = state("Height")
    End If
    layouts.Remove f.Name
End Sub
' Controls that use extra width: lists, subforms, wide text boxes and wide
' captions or outlines. Buttons, combo boxes and short inputs keep their size.
Private Function UXElastic(ByVal c As Access.Control) As Boolean
    Select Case c.ControlType
        Case acListBox, acSubform: UXElastic = True
        Case acTextBox: UXElastic = (c.Width >= 4000)
        Case acLabel, acRectangle: UXElastic = (c.Width >= 6000)
    End Select
End Function
' Share the extra width among the elastic controls beside v (sharing rows, not
' overlapping it horizontally), in proportion to their design widths. Controls to
' their right move over by the share they pass. Returns the offset from the rail.
Private Sub UXElasticPlace(ByVal geometry As Object, ByVal v As Variant, ByVal extra As Long, ByRef x As Long, ByRef w As Long)
    Dim key As Variant, p As Variant, total As Double, before As Double
    If v(4) Then total = v(2)
    For Each key In geometry.Keys
        p = geometry(key)
        If p(4) And p(1) < v(1) + v(3) And v(1) < p(1) + p(3) Then
            If p(0) + p(2) <= v(0) Then
                total = total + p(2): before = before + p(2)
            ElseIf p(0) >= v(0) + v(2) Then
                total = total + p(2)
            End If
        End If
    Next
    x = v(0) - 2850: w = v(2)
    If total > 0 Then
        x = x + CLng(extra * before / total)
        If v(4) Then w = w + CLng(extra * v(2) / total)
    End If
End Sub
Public Function UXResizeWorkspace(ByVal formName As String, Optional ByVal viewportWidth As Long = 0, Optional ByVal viewportHeight As Long = 0) As Boolean
    Dim f As Access.Form, c As Access.Control, state As Object, geometry As Object, old As Access.Form
    Dim width As Long, height As Long, left As Long, available As Long, baseWidth As Long, bottom As Long, ratio As Double, v As Variant, compact As Boolean
    Dim x As Long, w As Long, painted As Boolean, extra As Long
    If viewportBusy Then Exit Function
    On Error GoTo Failed
    Set f = Forms(formName)
    If Not UXHasControl(f, "uxCompactNav") Or formName = "frmEditor" Then Exit Function
    width = f.InsideWidth: height = f.InsideHeight
    If viewportWidth > 0 Then width = viewportWidth
    If viewportHeight > 0 Then height = viewportHeight
    If width < 4000 Or height < 2000 Then Exit Function
    If width > UX_MAX_FORM_WIDTH Then width = UX_MAX_FORM_WIDTH
    viewportBusy = True: viewportError = ""
    If layouts Is Nothing Then Set layouts = CreateObject("Scripting.Dictionary")
    If layouts.Exists(formName) Then
        Set state = layouts(formName): Set old = state("Form")
        If Not old Is f Then layouts.Remove formName
    End If
    If Not layouts.Exists(formName) Then
        Set state = CreateObject("Scripting.Dictionary"): Set geometry = CreateObject("Scripting.Dictionary")
        baseWidth = 10450
        For Each c In f.Controls
            If c.Left >= 2450 And c.Name <> "cmdWindowHelp" And c.Name <> "uxCompactNav" And c.Name <> "uxContext" And VBA.Left$(c.Tag, 5) <> "SORT:" Then
                geometry.Add c.Name, Array(c.Left, c.Top, c.Width, c.Height, UXElastic(c))
                If c.Visible And c.Left + c.Width - 2850 > baseWidth Then baseWidth = c.Left + c.Width - 2850
            End If
        Next
        state.Add "Form", f: state.Add "Geometry", geometry: state.Add "Width", baseWidth: state.Add "Height", f.Section(acDetail).Height
        layouts.Add formName, state
    End If
    Set state = layouts(formName): Set geometry = state("Geometry")
    If state.Exists("AppliedWidth") Then
        If state("AppliedWidth") = width And state("AppliedHeight") = height Then
            UXResizeWorkspace = True
            GoTo Finished
        End If
    End If
    compact = (width < 12000): left = IIf(compact, 250, 2850): available = width - left - 350
    ratio = available / state("Width")
    'Each geometry write repaints; hold painting until the whole pass is done.
    f.Painting = False: painted = True
    UXNavigationLayout f, compact, width
    UXVisualSet f!uxContext, "Left", IIf(compact, 2950, 2850): UXVisualSet f!uxContext, "Top", 200
    UXVisualSet f!uxContext, "Width", width - f!uxContext.Left - 350
    bottom = 0
    'Narrower than the design: scale everything. Wider: only elastic controls grow.
    If ratio > 1 Then extra = IIf(available > MAX_CONTENT_WIDTH, MAX_CONTENT_WIDTH, available) - state("Width")
    If extra < 0 Then extra = 0
    For Each c In f.Controls
        If geometry.Exists(c.Name) Then
            v = geometry(c.Name)
            If ratio > 1 Then
                UXElasticPlace geometry, v, extra, x, w: x = x + left
            Else
                x = left + CLng((v(0) - 2850) * ratio): w = CLng(v(2) * ratio)
                If w < 200 Then w = 200
            End If
            If x + w > width - 250 Then w = width - 250 - x
            UXVisualSet c, "Left", x: UXVisualSet c, "Width", w
            UXVisualSet c, "Top", v(1): UXVisualSet c, "Height", v(3)
            If c.Visible And c.Top + c.Height > bottom Then bottom = c.Top + c.Height
        End If
    Next
    'Keep authored row grouping and readable font sizes. Vertical scrolling is
    'the fallback for shorter windows; horizontal scrolling is never required.
    UXVisualSet f, "ScrollBars", 2: UXVisualSet f, "Width", width - 100
    UXPositionGroupFooter f
    If UXHasControl(f, "uxCurrentExport") Then
        If f!uxCurrentExport.Visible Then bottom = f!uxCurrentExport.Top + f!uxCurrentExport.Height
    End If
    UXVisualSet f!uxRail, "Height", bottom + 300
    UXVisualSet f.Section(acDetail), "Height", bottom + 300
    state("AppliedWidth") = width: state("AppliedHeight") = height
    layoutPasses = layoutPasses + 1
    UXResizeWorkspace = True
Finished:
    'A failure here must not resume into Finished again: that looped forever.
    On Error GoTo CleanupFailed
    If Not f Is Nothing Then
        UXSortHeaders f
        UXWindowHelpLayout f.Name, width
        UXApplyVisual f
        If painted Then f.Painting = True
    End If
    viewportBusy = False
    Exit Function
Failed:
    viewportError = "Viewport " & formName & ": " & Err.Number & " " & Err.Description
    Debug.Print viewportError
    Resume Finished
CleanupFailed:
    If Len(viewportError) = 0 Then viewportError = "Viewport cleanup " & formName & ": " & Err.Number & " " & Err.Description: Debug.Print viewportError
    Resume Next
End Function
