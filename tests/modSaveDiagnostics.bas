Attribute VB_Name = "modSaveDiagnostics"
Option Explicit
Public Sub SaveTrace(ByVal message As String)
 Dim n As Integer
 On Error Resume Next
 n = FreeFile
 Open CurrentProject.FullName & ".save-trace.log" For Append As #n
 Print #n, Format$(Now(), "yyyy-mm-dd hh:nn:ss") & " " & message
 Close #n
End Sub
