Attribute VB_Name = "ReviewHarness"
Option Explicit
Public Function TryReview(ByVal path As String) As String
    On Error GoTo Failed
    If OpenReviewFile(path) Then TryReview = "PASS"
    Exit Function
Failed:
    TryReview = CStr(Err.Number) & ": " & Err.Description
End Function
