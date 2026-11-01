Attribute VB_Name = "modUXViewport"
Option Explicit
Private layouts As Object
Private viewportBusy As Boolean
Private viewportError As String
Private layoutPasses As Long
' Workspace geometry, in twips. Controls left of UX_RAIL_EDGE belong to the
' navigation rail; content starts at UX_CONTENT_LEFT (UX_COMPACT_LEFT when the
' window is narrower than UX_COMPACT_WIDTH and the rail becomes a drop-down).
' Content is designed UX_DESIGN_WIDTH wide and stretches to MAX_CONTENT_WIDTH;
' wider windows leave a right margin instead of over-long rows.
Public Const UX_RAIL_EDGE As Long = 2450
Public Const UX_CONTENT_LEFT As Long = 2850
Public Const UX_COMPACT_LEFT As Long = 250
Public Const UX_COMPACT_WIDTH As Long = 12000
Public Const UX_DESIGN_WIDTH As Long = 10450
Private Const UX_RIGHT_MARGIN As Long = 350
Private Const MAX_CONTENT_WIDTH As Long = 16500
' How far below a fixed-width label to look for the field it captions.
Private Const CAPTION_REACH As Long = 150
' The group members grid shows twelve rows by design and grows by whole rows
' (frmUXTraitRows detail height) into spare window height, up to fifteen more.
' The group footer (senior caption and buttons) sits 930 twips under content.
Private Const TALL_ROW_HEIGHT As Long = 430
Private Const TALL_MAX_GROWTH As Long = 6450
Private Const GROUP_FOOTER_HEIGHT As Long = 930
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
                v = geometry(c.Name)
                If c.ControlType = acListBox And c.Width <> v(2) Then UXScaleColumns c, c.Width, v(2)
                c.Left = v(0): c.Top = v(1): c.Width = v(2): c.Height = v(3)
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
' their right move over by the share they pass. A fixed-width label directly above
' a control with the same left edge is its caption and moves with it.
' Returns the offset from the rail.
Private Sub UXElasticPlace(ByVal geometry As Object, ByVal v As Variant, ByVal extra As Long, ByRef x As Long, ByRef w As Long)
    Dim key As Variant, p As Variant, total As Double, before As Double, captioned As Long, slots As Object
    If v(5) And Not v(4) Then
        For Each key In geometry.Keys
            p = geometry(key)
            If p(0) = v(0) And p(1) >= v(1) + v(3) And p(1) <= v(1) + v(3) + CAPTION_REACH Then
                UXElasticPlace geometry, p, extra, x, captioned: w = v(2): Exit Sub
            End If
        Next
    End If
    If v(4) Then total = v(2)
    'Controls stacked in one slot (other pages or views) count once, so hidden
    'alternatives never shrink a row's share.
    Set slots = CreateObject("Scripting.Dictionary")
    For Each key In geometry.Keys
        p = geometry(key)
        If p(4) And p(1) < v(1) + v(3) And v(1) < p(1) + p(3) And Not slots.Exists(p(0) & ":" & p(2)) Then
            slots.Add p(0) & ":" & p(2), True
            If p(0) + p(2) <= v(0) Then
                total = total + p(2): before = before + p(2)
            ElseIf p(0) >= v(0) + v(2) Then
                total = total + p(2)
            End If
        End If
    Next
    x = v(0) - UX_CONTENT_LEFT: w = v(2)
    If total > 0 Then
        x = x + CLng(extra * before / total)
        If v(4) Then w = w + CLng(extra * v(2) / total)
    End If
End Sub
' The list or grid on a screen that takes spare window height: the group
' members grid, or the staged-changes list that replaces it in preview.
Private Function UXTallControl(ByVal f As Access.Form) As String
    If f.Name = "frmUXComparison" Then UXTallControl = IIf(f!uxTraitGrid.Visible, "uxTraitGrid", "lstProposals")
End Function
' Height the tall control gains: the window height left under the visible
' content and the group footer, in whole grid rows, up to TALL_MAX_GROWTH.
Private Function UXTallGrowth(ByVal f As Access.Form, ByVal geometry As Object, ByVal height As Long) As Long
    Dim key As Variant, p As Variant, c As Access.Control, content As Long, spare As Long
    For Each key In geometry.Keys
        Set c = f.Controls(key): p = geometry(key)
        If c.Visible And c.Tag <> "GroupFooter" And p(1) + p(3) > content Then content = p(1) + p(3)
    Next
    spare = height - content - IIf(f!uxCurrentExport.Visible, GROUP_FOOTER_HEIGHT, 0) - 300
    If spare <= 0 Then Exit Function
    spare = (spare \ TALL_ROW_HEIGHT) * TALL_ROW_HEIGHT
    UXTallGrowth = IIf(spare > TALL_MAX_GROWTH, TALL_MAX_GROWTH, spare)
End Function
' Keep a list's columns in proportion when the list changes width. Screens set
' their own column mixes (for example the Reports column views), so scale the
' current widths rather than restoring design ones.
Private Sub UXScaleColumns(ByVal c As Access.Control, ByVal oldWidth As Long, ByVal newWidth As Long)
    Dim parts As Variant, i As Long, sizes As String
    If oldWidth <= 0 Or Len(c.ColumnWidths) = 0 Then Exit Sub
    parts = Split(c.ColumnWidths, ";")
    For i = 0 To UBound(parts)
        If i > 0 Then sizes = sizes & ";"
        If Len(parts(i)) > 0 Then sizes = sizes & CLng(Val(parts(i)) * newWidth / oldWidth)
    Next
    c.ColumnWidths = sizes
End Sub
Public Function UXResizeWorkspace(ByVal formName As String, Optional ByVal viewportWidth As Long = 0, Optional ByVal viewportHeight As Long = 0) As Boolean
    Dim f As Access.Form, c As Access.Control, state As Object, geometry As Object, old As Access.Form
    Dim width As Long, height As Long, left As Long, available As Long, baseWidth As Long, bottom As Long, ratio As Double, v As Variant, compact As Boolean
    Dim x As Long, w As Long, painted As Boolean, extra As Long, oldWidth As Long, tall As String, growth As Long, tallBottom As Long
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
        baseWidth = UX_DESIGN_WIDTH
        For Each c In f.Controls
            If c.Left >= UX_RAIL_EDGE And c.Name <> "cmdWindowHelp" And c.Name <> "uxCompactNav" And c.Name <> "uxContext" And VBA.Left$(c.Tag, 5) <> "SORT:" Then
                geometry.Add c.Name, Array(c.Left, c.Top, c.Width, c.Height, UXElastic(c), c.ControlType = acLabel)
                If c.Visible And c.Left + c.Width - UX_CONTENT_LEFT > baseWidth Then baseWidth = c.Left + c.Width - UX_CONTENT_LEFT
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
    compact = (width < UX_COMPACT_WIDTH): left = IIf(compact, UX_COMPACT_LEFT, UX_CONTENT_LEFT): available = width - left - UX_RIGHT_MARGIN
    ratio = available / state("Width")
    'Each geometry write repaints; hold painting until the whole pass is done.
    f.Painting = False: painted = True
    UXNavigationLayout f, compact, width
    'Compact: the context line follows the 2,500-twip navigation drop-down.
    UXVisualSet f!uxContext, "Left", IIf(compact, UX_COMPACT_LEFT + 2700, UX_CONTENT_LEFT): UXVisualSet f!uxContext, "Top", 200
    UXVisualSet f!uxContext, "Width", width - f!uxContext.Left - UX_RIGHT_MARGIN
    bottom = 0
    'Narrower than the design: scale everything. Wider: only elastic controls grow.
    If ratio > 1 Then extra = IIf(available > MAX_CONTENT_WIDTH, MAX_CONTENT_WIDTH, available) - state("Width")
    If extra < 0 Then extra = 0
    tall = UXTallControl(f)
    If Len(tall) > 0 Then v = geometry(tall): tallBottom = v(1) + v(3): growth = UXTallGrowth(f, geometry, height)
    For Each c In f.Controls
        If geometry.Exists(c.Name) Then
            v = geometry(c.Name)
            If ratio > 1 Then
                UXElasticPlace geometry, v, extra, x, w: x = x + left
            Else
                x = left + CLng((v(0) - UX_CONTENT_LEFT) * ratio): w = CLng(v(2) * ratio)
                If w < 200 Then w = 200
            End If
            If x + w > width - 250 Then w = width - 250 - x
            'Scale list columns by the width Access actually applied, so a refused
            'write never compounds the columns on the next pass.
            oldWidth = c.Width
            UXVisualSet c, "Left", x: UXVisualSet c, "Width", w
            If c.ControlType = acListBox And c.Width <> oldWidth Then UXScaleColumns c, oldWidth, c.Width
            If c.Name = tall Then
                UXVisualSet c, "Top", v(1): UXVisualSet c, "Height", v(3) + growth
            ElseIf growth > 0 And v(1) >= tallBottom Then
                UXVisualSet c, "Top", v(1) + growth: UXVisualSet c, "Height", v(3)
            Else
                UXVisualSet c, "Top", v(1): UXVisualSet c, "Height", v(3)
            End If
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
    'The rail runs the full window height even when the content is shorter.
    bottom = bottom + 300: If bottom < height - 15 Then bottom = height - 15
    UXVisualSet f!uxRail, "Height", bottom
    UXVisualSet f.Section(acDetail), "Height", bottom
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
