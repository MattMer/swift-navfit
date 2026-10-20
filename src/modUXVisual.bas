Attribute VB_Name = "modUXVisual"
Option Explicit
Private painting As Boolean
Public Function UXColor(ByVal role As String) As Long
 Select Case role
  Case "Canvas": UXColor = RGB(244, 247, 251)
  Case "Ink": UXColor = RGB(30, 45, 64)
  Case "Rail": UXColor = RGB(20, 45, 78)
  Case "Primary": UXColor = RGB(36, 91, 143)
  Case "Active": UXColor = RGB(14, 112, 107)
  Case "Neutral": UXColor = RGB(229, 236, 243)
  Case "Border": UXColor = RGB(167, 184, 202)
  Case "Warning": UXColor = RGB(255, 240, 207)
  Case "WarningInk": UXColor = RGB(122, 71, 0)
  Case "Error": UXColor = RGB(252, 232, 232)
  Case "ErrorInk": UXColor = RGB(161, 38, 38)
  Case "Success": UXColor = RGB(226, 242, 233)
  Case "SuccessInk": UXColor = RGB(31, 101, 68)
  Case "Help": UXColor = RGB(238, 232, 250)
  Case "HelpInk": UXColor = RGB(99, 61, 136)
  Case "Muted": UXColor = RGB(100, 116, 133)
  Case Else: UXColor = RGB(255, 255, 255)
 End Select
End Function
Public Function UXVisualEvent(ByVal formName As String, ByVal original As String) As Boolean
 If Len(original) > 0 Then Eval original
 UXVisualEvent = UXApplyVisual(Forms(formName))
End Function
Public Function UXApplyVisual(ByVal f As Access.Form) As Boolean
 Dim c As Access.Control, active As Long, key As String, text As String, role As String, tone As String
 If painting Then Exit Function
 painting = True
 On Error GoTo Finished
 UXVisualSet f.Section(acDetail), "BackColor", UXColor("Canvas")
 Select Case f.Name
  Case "frmLibrary", "frmEditor": active = 1
  Case "frmUXGroups", "frmMajorWorkbench", "frmUXComparison", "frmUXBulk", "frmUXGroupRSCA", "frmUXNewGroup", "frmUXGroupReport": active = 2
  Case "frmMyWork", "frmUXReview", "frmReviewComments": active = 3
  Case "frmUXResources", "frmUXSeniors", "frmUXSeniorEditor", "frmUXDirectory", "frmUXNavyWriter", "frmSnippets": active = 4
  Case "frmUXSettings", "frmBackup", "frmRecovery", "frmOperations": active = 5
 End Select
 For Each c In f.Controls
  key = LCase$(c.Name)
  Select Case c.ControlType
   Case acCommandButton
    UXVisualSet c, "UseTheme", True: UXVisualSet c, "BorderColor", UXColor("Border")
    role = "Neutral": tone = "Ink"
    If key Like "uxnav#" Then
     UXVisualSet c, "FontBold", (Val(Right$(key, 1)) = active)
     role = "Rail": tone = "White"
     If c.FontBold Then role = "Active"
    ElseIf key = "cmdwindowhelp" Or InStr(key, "help") > 0 Or key = "cmdworkspaceguide" Then
     role = "Help": tone = "HelpInk"
    ElseIf c.FontBold And (InStr(key, "tab") > 0 Or InStr(key, "page") > 0 Or key Like "uxsection#") Then
     role = "Active": tone = "White"
    ElseIf key Like "*archive*" Or key Like "*discard*" Or key Like "*delete*" Or key Like "*return*" Then
     role = "Error": tone = "ErrorInk"
    ElseIf c.FontBold Or key = "cmdsave" Or key = "cmdapply" Or key = "cmdcreate" Or key = "cmdnew" Or key = "cmdimport" Or key = "cmdexport" Or key = "cmdopen" Then
     role = "Primary": tone = "White"
    End If
    If key = "uxdirty" Then
     role = UXStatusRole(c.Caption, key)
     If Len(role) = 0 Then role = "Neutral"
     tone = IIf(role = "Neutral", "Ink", role & "Ink")
    End If
    If Not c.Enabled Then role = "Neutral": tone = "Muted"
    UXVisualSet c, "BackColor", UXColor(role): UXVisualSet c, "ForeColor", UXColor(tone)
    UXVisualSet c, "HoverColor", UXColor(role): UXVisualSet c, "PressedColor", UXColor(role)
    UXVisualSet c, "HoverForeColor", UXColor(tone): UXVisualSet c, "PressedForeColor", UXColor(tone)
   Case acLabel
    UXVisualSet c, "ForeColor", UXColor("Ink")
    If key = "uxbrand" Or key = "uxnavcaption" Then UXVisualSet c, "ForeColor", RGB(255, 255, 255)
    If key = "uxerror" Or key Like "err*" Or key = "uxdirty" Or key = "lblsaved" Or key = "lblrecovery" Or key Like "*status" Or key = "uxscope" Or key = "uxlocked" Then
     role = UXStatusRole(c.Caption, key): tone = "Ink"
     If Len(role) > 0 Then tone = role & "Ink"
     UXVisualSet c, "BackStyle", IIf(Len(role) > 0, 1, 0)
     If Len(role) > 0 Then UXVisualSet c, "BackColor", UXColor(role)
     UXVisualSet c, "ForeColor", UXColor(tone)
    End If
   Case acRectangle
    If key = "uxrail" Then UXVisualSet c, "BackStyle", 1: UXVisualSet c, "BackColor", UXColor("Rail"): UXVisualSet c, "BorderColor", UXColor("Rail")
   Case acTextBox, acComboBox, acListBox
    If f.Name = "frmUXTraitRows" And (key = "priorita" Or key = "priorpromotion") Then
     UXVisualSet c, "BackColor", UXColor("Neutral"): UXVisualSet c, "ForeColor", UXColor("Muted"): UXVisualSet c, "BorderColor", UXColor("Border")
    ElseIf f.CurrentView = 0 Or c.ControlType = acListBox Then
     UXVisualSet c, "BackColor", RGB(255, 255, 255): UXVisualSet c, "ForeColor", UXColor("Ink"): UXVisualSet c, "BorderColor", UXColor("Border")
    End If
  End Select
 Next
 If f.Name = "frmEditor" Then
  If Left$(f!uxITA.Caption, 14) = "ITA - pending:" Then
   f!uxITA.ControlTipText = Left$(f!uxITA.Caption, 255)
   f!uxITA.Caption = Replace(Split(f!uxITA.Caption, ";")(0), "ITA - pending: Numeric average:", "ITA (pending):")
  End If
  If Left$(f!uxSGA.Caption, Len("Summary Group Average - saved values:")) = "Summary Group Average - saved values:" Then
   f!uxSGA.ControlTipText = Left$(f!uxSGA.Caption, 255)
   f!uxSGA.Caption = Replace(Split(f!uxSGA.Caption, " (")(0), "Summary Group Average - saved values:", "Group average (saved):")
  End If
 End If
 UXApplyVisual = True
Finished:
 painting = False
End Function

Public Function UXVisualRailBack(ByVal formName As String) As Boolean
 Dim f As Access.Form, c As Access.Control
 Set f = Forms(formName)
 If Not UXHasControl(f, "uxRail") Then UXVisualRailBack = True: Exit Function
 For Each c In f.Controls: c.InSelection = False: Next
 f!uxRail.InSelection = True
 DoCmd.RunCommand acCmdSendToBack
 f!uxRail.InSelection = False
 UXVisualRailBack = True
End Function
Public Function UXStatusRole(ByVal caption As String, ByVal key As String) As String
 Dim text As String
 text = LCase$(caption)
 If Len(text) = 0 Then Exit Function
 If InStr(text, "no unsaved edits") > 0 Or Left$(text, 15) = "review updated:" Then
  UXStatusRole = "Success"
 ElseIf key = "uxerror" Or key Like "err*" Or InStr(text, "failed") > 0 Or InStr(text, "not saved") > 0 Or InStr(text, "unavailable") > 0 Then
  UXStatusRole = "Error"
 ElseIf InStr(text, "unsaved") > 0 Or InStr(text, "pending edits") > 0 Or (InStr(text, "issues") > 0 And InStr(text, "0 issues") = 0) Or InStr(text, "scope:") > 0 Or InStr(text, "stale") > 0 Then
  UXStatusRole = "Warning"
 ElseIf InStr(text, "saved") > 0 Or InStr(text, "no pending") > 0 Or InStr(text, "approved") > 0 Then
  UXStatusRole = "Success"
 End If
End Function

' Avoid redundant Access property writes and their invalidation/repaint cost.
Private Sub UXVisualSet(ByVal target As Object, ByVal property As String, ByVal value As Variant)
 If target.Properties(property).Value <> value Then target.Properties(property).Value = value
End Sub
