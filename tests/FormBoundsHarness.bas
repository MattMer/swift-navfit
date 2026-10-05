Attribute VB_Name = "FormBoundsHarness"
Option Explicit
Public Function TestFormBounds(ByVal name As String) As Boolean
    Dim f As Access.Form, c As Access.Control
    Set f = Forms(name)
    For Each c In f.Controls
        If c.Visible Then
            If c.Left < 0 Or c.Left + c.Width > f.Width Or c.Top < 0 Or c.Top + c.Height > f.Section(acDetail).Height Then Err.Raise vbObjectError + 5400, , "Control outside form: " & name & "/" & c.Name
        End If
    Next
    TestFormBounds = True
End Function
Public Function TestButtonOverlap(ByVal name As String, ByVal buttonName As String) As Boolean
    Dim f As Access.Form, button As Access.Control, c As Access.Control
    Set f = Forms(name): Set button = f.Controls(buttonName)
    For Each c In f.Controls
        If button.Visible And c.Visible And c.ControlType <> acRectangle And c.Name <> button.Name Then
            If button.Left < c.Left + c.Width And button.Left + button.Width > c.Left And button.Top < c.Top + c.Height And button.Top + button.Height > c.Top Then Err.Raise vbObjectError + 5401, , "Button overlaps " & name & "/" & c.Name
        End If
    Next
    TestButtonOverlap = True
End Function
