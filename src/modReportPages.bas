Attribute VB_Name = "modReportPages"
Option Explicit
Private sheet As Long, renders As Long, arranging As Boolean
Private Const ReportLeft As Long = 2600
Public Function UXFixedEnabled() As Boolean
 If CurrentProject.AllForms("frmEditor").IsLoaded Then UXFixedEnabled = UXHasControl(Forms!frmEditor, "uxFrontPage")
End Function
Public Function UXFixedRenderCount() As Long
 UXFixedRenderCount = renders
End Function
Private Sub Place(ByVal f As Access.Form, ByVal name As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
 If f.Name = "frmEditor" Then x = x + ReportLeft
 With f.Controls(name)
  If .Left <> x Then .Left = x
  If .Top <> y Then .Top = y
  If .Width <> width Then .Width = width
  If .Height <> height Then .Height = height
  If Not .Visible Then .Visible = True
 End With
End Sub
'Fixed sheet geometry never depends on the window. Only the scroll viewport and
'action bar change height; resizing does not rebuild fields or refresh data.
Public Function UXFixedResize(Optional ByVal height As Long = 0, Optional ByVal viewportWidth As Long = 0) As Boolean
 Dim f As Access.Form, body As Long, i As Long, names As Variant, c As Access.Control, width As Long, viewingWidth As Long
 If arranging Then Exit Function
 arranging = True
 On Error GoTo Failed
 Set f = Forms!frmEditor
 If height = 0 Then height = f.InsideHeight
 If height < 7000 Then height = 7000
 body = height - 5000
 width = f.InsideWidth
 If viewportWidth > 0 Then width = viewportWidth
 viewingWidth = width - (250 + ReportLeft) - 400
 If viewingWidth < 12600 Then viewingWidth = 12600
 f.ScrollBars = 3
 If f.Width <> viewingWidth + ReportLeft + 650 Then f.Width = viewingWidth + ReportLeft + 650
 'The fixed paper canvas lives beside the application rail, never over it.
 UXNavigationLayout f, False, f.Width
 Place f, "uxContext", 250, 100, 8500, 300
 Place f, "uxReport", 250, 550, 8700, 650
 Place f, "uxTypeCaption", 9500, 500, 1200, 300
 If UXHasControl(f, "uxCreatePDF") Then Place f, "uxCreatePDF", 10800, 450, 1200, 400
 Place f, "cboKind", 9500, 850, 2500, 350
 Place f, "uxITA", 250, 1250, 5500, 350
 Place f, "uxSGA", 5800, 1250, 6200, 350
 Place f, "uxLocked", 250, 1630, 8700, 350
 Place f, "uxAlias", 250, 1980, 8700, 300
 Place f, "uxGroup", 9200, 1750, 2800, 400
 names = Array("uxReviewSetup", "uxReviewAccept", "uxReviewReturn", "uxReviewApproval")
 For i = 0 To 3: Place f, CStr(names(i)), 250 + i * 3000, 2350, 2850, 400: Next
 names = Array("uxFrontPage", "uxBackPage", "uxPageReview", "uxPageHistory")
 For i = 0 To 3: Place f, CStr(names(i)), 250 + i * 3000, 2850, 2850, 400: Next
 Place f, "uxSection", 250, 3320, 11800, 350
 Place f, "uxError", 250, 3690, 11800, 350
 f!uxErrorList.Left = 250 + ReportLeft: f!uxErrorList.Top = 4060: f!uxErrorList.Width = 11800: f!uxErrorList.Height = 500
 Place f, "uxFieldSurface", 250, IIf(f!uxErrorList.Visible, 4620, 4060), viewingWidth, body
 f!uxFieldSurface.Visible = (sheet < 8)
 Dim footer As Long
 footer = f!uxFieldSurface.Top + body + 120
 Place f, "lblSaved", 250, footer, 5600, 300
 Place f, "uxDirty", 6000, footer, 6000, 300
 Place f, "cmdClose", 250, footer + 370, 2250, 450
 Place f, "uxUndoPending", 2700, footer + 370, 3000, 450
 Place f, "uxPageValidation", 5950, footer + 370, 2700, 450
 Place f, "cmdSave", 9000, footer + 370, 3000, 450
 Place f, "lblRecovery", 250, footer + 860, 11750, 300
 If UXHasControl(f, "uxToolsSurface") Then
  With f!uxToolsSurface
   .Left = 250 + ReportLeft: .Top = 4800: .Width = viewingWidth: .Height = body - 740
   If .Height < 2400 Then .Height = 2400
   .Visible = (sheet = 8)
  End With
 End If
 If sheet = 8 Then
  names = Array("cmdComments", "cmdRoute", "cmdChecklist", "cmdCalculations")
  For i = 0 To 3: Place f, CStr(names(i)), 250 + i * 3000, 4150, 2850, 450: Next
  f!uxReviewCommentGuide.Visible = False
 ElseIf sheet = 9 Then
  Place f, "cmdCompare", 250, 4150, 2800, 450: Place f, "cmdRestore", 3250, 4150, 2800, 450
  Place f, "lstRevisions", 250, 4750, 11800, body - 700
 End If
 If sheet = 0 Then
  Place f, "uxSeniorRecord", 250, footer + 1250, 7600, 400
  Place f, "uxUseSenior", 8100, footer + 1250, 3900, 400
 End If
 f.Section(acDetail).Height = footer + IIf(sheet = 0, 1850, 1250)
 f!uxRail.Height = f.Section(acDetail).Height
 arranging = False: UXFixedResize = True: Exit Function
Failed:
 arranging = False: f!uxError.Caption = "Layout: " & Err.Description
End Function
Public Function UXFixedSelect(ByVal section As Long) As Boolean
 Dim f As Access.Form, surface As Access.Form, c As Access.Control, r As DAO.Recordset, db As DAO.Database, locked As Boolean, id As Long, parentPaint As Boolean, surfacePaint As Boolean, paintHeld As Boolean
 On Error GoTo Failed
 If Not UXToolsCheckpoint() Then Exit Function
 UXFlushSurface
 Set f = Forms!frmEditor: Set surface = f!uxFieldSurface.Form
 parentPaint = f.Painting: surfacePaint = surface.Painting
 f!cmdClose.SetFocus
 If section <= 4 Then
  sheet = 0
 ElseIf section <= 6 Then
  sheet = 1
 ElseIf section = 7 Then
  UXToolsSetNotes
  sheet = 8
 Else
  sheet = section
 End If
 'Hide the native child window before suspending parent painting. Otherwise
 'Access can leave its last notes page covering the parent's Review actions.
 If sheet >= 8 Then
  surface!uxFocusAnchor.SetFocus
  f!cmdClose.SetFocus
  For Each c In surface.Controls
   If c.Name <> "uxFocusAnchor" Then c.Visible = False
  Next
  f!uxFieldSurface.Visible = False
 End If
 f.Painting = False: surface.Painting = False: paintHeld = True
 id = Nz(f!txtReportID, 0): Set db = CurrentDb
 If id > 0 Then
  Set r = db.OpenRecordset("SELECT IsArchived,WorkflowState FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
  locked = r!IsArchived Or r!WorkflowState <> "Draft": r.Close
 End If
 For Each c In f.Controls
  If c.Name Like "uxSection#" Then c.Visible = False
  If Left$(c.Tag, 6) = "LOCAL:" Then c.Visible = (CLng(Mid$(c.Tag, 7)) = sheet)
 Next
 f!cmdSave.Enabled = Not locked: f!cboKind.Locked = locked
 f!uxLocked.Caption = IIf(locked, "Read only. Reopen local approval or recover the report to edit.", "")
 Select Case sheet
  Case 0: f!uxSection.Caption = "Front Page"
  Case 1: f!uxSection.Caption = "Back Page"
  Case 7: f!uxSection.Caption = "Report notes (local)"
  Case 8: f!uxSection.Caption = "Report tools"
  Case 9: f!uxSection.Caption = "History (local)"
 End Select
 If sheet < 2 Then
  UXRenderFixedSheet sheet, locked
 ElseIf sheet = 7 Then
  UXSurfaceSection 7, locked
 End If
 f!uxFrontPage.FontBold = (sheet = 0): f!uxBackPage.FontBold = (sheet = 1)
 UXFixedResize
 If sheet = 8 Then UXToolsRefresh
 If sheet = 7 Then UXSortHeaders Forms!frmEditor!uxFieldSurface.Form
 UXRefreshEditor
 UXHeaderReviewState
 f!uxFieldSurface.Visible = (sheet < 8)
 UXFixedSelect = True
Finished:
 If paintHeld Then
  surface.Painting = surfacePaint: f.Painting = parentPaint
  If parentPaint Then f.Repaint
 End If
 Exit Function
Failed:
 On Error Resume Next
 f!uxError.Caption = "Unable to open report page: " & Err.Number & " " & Err.Description
 UXFixedSelect = False
 Resume Finished
End Function
Private Sub UXRenderFixedSheet(ByVal page As Long, ByVal locked As Boolean)
 Dim f As Access.Form, row As Variant, parts As Variant, key As String, kind As String, c As Access.Control, y As Long, bottom As Long, trait As Boolean, order As Long
 Set f = Forms!frmEditor!uxFieldSurface.Form: kind = Nz(Forms!frmEditor!cboKind, "EVAL")
 f!uxFocusAnchor.SetFocus
 For Each c In f.Controls
  If c.Name <> "uxFocusAnchor" And c.Visible Then c.Visible = False
 Next
 renders = renders + 1
 For Each row In ReportPageRows(kind, page)
  parts = Split(CStr(row), "|"): key = parts(0): y = CLng(parts(2))
  If key = "@guidance" Then
   If f.Section(acDetail).Height < y + CLng(parts(4)) + 200 Then f.Section(acDetail).Height = y + CLng(parts(4)) + 200
   Place f, "uxTraitGuidance", CLng(parts(1)), y, CLng(parts(3)), CLng(parts(4))
  Else
   trait = IsTraitKey(key)
   Place f, "label" & key, CLng(parts(1)), y, CLng(parts(3)), 340
   f.Controls("label" & key).Caption = UXFieldCaption(key, kind): f.Controls("label" & key).FontSize = 8
   Place f, UXFieldControl(key), CLng(parts(1)), y + 350, CLng(parts(3)), CLng(parts(4))
   Set c = f.Controls(UXFieldControl(key)): c.TabIndex = order: order = order + 1
   c.Locked = locked Or Not UXFieldApplies(key, kind)
   If c.ControlType <> acCheckBox Then c.FontName = "Segoe UI": c.FontSize = 10
   If UXFieldFlag(key) Then
    c.Left = CLng(parts(1)): c.Top = y + 40: c.Width = 280: c.Height = 280
    With f.Controls("label" & key): .Left = c.Left + 350: .Width = CLng(parts(3)) - 350: .Height = 430: End With
   End If
   If trait Then
    f.Controls("label" & key).Width = 4000
    c.Left = 4400: c.Top = y: c.Width = 1000
    If kind = "CHIEFEVAL" Then
     f.Controls("label" & key).Width = 3300
     c.Left = 3600
     Place f, "uxChiefGuidance" & Mid$(key, 6), 180, y + 430, 4400, 1900
     f.Controls("uxChiefGuidance" & Mid$(key, 6)).Caption = CatalogueTraitDescription(kind, CLng(Mid$(key, 6)), Nz(Forms!frmEditor.Controls("cbo" & key), ""))
     f.Controls("uxChiefGuidance" & Mid$(key, 6)).FontName = "Times New Roman"
     f.Controls("uxChiefGuidance" & Mid$(key, 6)).FontSize = 8
    End If
   End If
   If key Like "Trait##Comment#" Then
    f.Controls("label" & key).Caption = "Comment " & Right$(key, 1)
    f.Controls("label" & key).Height = 190
    c.Top = y + 200: c.FontName = "Times New Roman": c.FontSize = 12
    c.LeftMargin = 0: c.RightMargin = 0: c.TopMargin = 0: c.BottomMargin = 0
    c.EnterKeyBehavior = True: c.ScrollBars = 2
    c.ControlTipText = UXFieldCaption("Trait" & Mid$(key, 6, 2), kind) & " / authored performance comment " & Right$(key, 1)
   End If
   If key = "OfficialComments" Then
    c.FontName = "Courier New": c.FontSize = UXNarrativeFontSize(kind): c.TextFormat = 0: c.Width = UXNarrativeWidth(kind)
    c.LeftMargin = 90: c.RightMargin = 90: c.TopMargin = 90: c.BottomMargin = 90: c.ScrollBars = 2
    Place f, "uxNarrativeCount", c.Left, c.Top + c.Height + 260, 11400, 400
    Place f, "cmdSnippets", 180, c.Top + c.Height + 700, 3200, 400: f!cmdSnippets.Enabled = Not locked
   End If
   Place f, "err" & key, CLng(parts(1)), y + IIf(trait, 390, IIf(UXFieldFlag(key), 430, CLng(parts(4)) + 360)), CLng(parts(3)), 200
   f.Controls("err" & key).FontSize = 8
   UXSurfaceValidate key
   bottom = y + CLng(parts(4)) + IIf(key = "OfficialComments", 1450, 650)
   If f.Section(acDetail).Height < bottom Then f.Section(acDetail).Height = bottom
  End If
 Next
 f.Width = 12200: f.ScrollBars = 3
 If f!uxTraitGuidance.Visible Then
  For Each row In ReportPageRows(kind, page)
   key = Split(CStr(row), "|")(0)
   If IsTraitKey(key) Then UXSurfaceTrait CLng(Mid$(key, 6)): Exit For
  Next
 End If
End Sub
'Used by validation findings and legacy detail actions to select the actual page.
Public Function UXFixedOpenField(ByVal key As String) As Boolean
 Dim page As Long
 page = ReportPageForField(Nz(Forms!frmEditor!cboKind, "EVAL"), key)
 UXFixedOpenField = UXSelectSection(IIf(page = 0, 0, 5))
End Function
