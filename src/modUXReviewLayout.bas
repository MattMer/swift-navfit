Attribute VB_Name = "modUXReviewLayout"
Option Explicit
Private busy As Boolean
Public Function UXReviewLayout(Optional ByVal viewportWidth As Long = 0) As Boolean
    Dim f As Access.Form, c As Access.Control, parts As Variant, width As Long, ratio As Double, offset As Long, bottom As Long
    If busy Then Exit Function
    On Error GoTo Finished
    Set f = Forms!frmUXReview: width = f.InsideWidth
    If viewportWidth > 0 Then width = viewportWidth
    If width < 5000 Then Exit Function
    busy = True: ratio = (width - 800) / 12600
    If width < 11000 Then offset = 450
    For Each c In f.Controls
        If Left$(c.Tag, 3) = "RV:" Then
            parts = Split(c.Tag, ":")
            c.Left = 400 + CLng((CLng(parts(1)) - 400) * ratio)
            c.Width = CLng(CLng(parts(2)) * ratio)
            c.Top = CLng(parts(3)): c.Height = CLng(parts(4))
            If c.Name = "lblStatus" Then c.Height = c.Height + offset
            If c.Top >= 4900 Then c.Top = c.Top + offset
            If c.Visible And c.Top + c.Height > bottom Then bottom = c.Top + c.Height
        End If
    Next
    f!lstSteps.ColumnWidths = CStr(CLng(650 * ratio)) & ";" & CStr(CLng(3300 * ratio)) & ";" & CStr(CLng(1800 * ratio)) & ";0;0"
    f.Width = width - 100: f.Section(acDetail).Height = bottom + 300: f.ScrollBars = 2
    UXReviewLayout = True
Finished:
    busy = False
End Function
