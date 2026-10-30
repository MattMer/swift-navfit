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
