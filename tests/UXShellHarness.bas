Attribute VB_Name = "UXShellHarness"
Option Explicit
Public Function TestUXShell() As String
    Dim name As Variant, f As Access.Form, c As Access.Control, checks As Long
    On Error GoTo Failed
    For Each name In Array("frmHome", "frmLibrary", "frmEditor", "frmUXGroups", "frmMyWork", "frmUXResources", "frmUXSettings", "frmMajorWorkbench", "frmUXBulk", "frmUXComparison", "frmUXSeniors", "frmUXSeniorEditor", "frmUXGroupRSCA")
        DoCmd.OpenForm CStr(name), acDesign: Set f = Forms(CStr(name))
        For Each c In f.Controls
            If c.Name Like "uxNav#" Then
                If c.Left >= 2450 Or Not c.Visible Then Err.Raise 5, , name & " navigation is unavailable"
                checks = checks + 1
            ElseIf c.Visible And c.ControlType <> acRectangle And c.Name <> "uxBrand" And c.Name <> "uxNavCaption" Then
                If c.Left < 2450 Or c.Left + c.Width > f.Width Then Err.Raise 5, , name & "/" & c.Name & " exceeds its content area"
                checks = checks + 1
            End If
        Next
        If name = "frmEditor" Then
            If Not f!cboKind.Visible Then Err.Raise 5, , "Report type selector is hidden"
        ElseIf name = "frmLibrary" Then
            If f!lstFolders.Visible Then Err.Raise 5, , "Legacy folder tree still visible"
        End If
        DoCmd.Close acForm, CStr(name), acSaveNo
    Next
    TestUXShell = "PASS: " & checks & " native navigation/content bounds checks; compiled successfully. Visual and assistive-technology testing still required.": Exit Function
Failed:
    TestUXShell = "FAIL: " & Err.Description
End Function
