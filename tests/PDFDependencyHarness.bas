Attribute VB_Name = "PDFDependencyHarness"
Option Compare Binary
Option Explicit
Public Function PDFTargetContext() As String
    Dim architecture As String
#If Win64 Then
    architecture = "64-bit"
#Else
    architecture = "32-bit"
#End If
    PDFTargetContext = "Access " & Application.Version & " build " & Application.Build & " " & architecture
End Function
Public Function PDFTargetProbe(ByVal path As String, ByVal kind As String) As String
    Dim p As Object, values As Object, before As String
    On Error GoTo Failed
    before = SourceFileHash(path)
    Set p = PreviewStructuredPDF(path)
    If p("FormKind") <> kind Then Err.Raise 5, , "Wrong template identity."
    Set values = PDFMappedValues("20", "SSN")
    If values.Count <> PDFMappingCount(kind) Then Err.Raise 5, , "Expanded synthetic field count differs."
    If values("UnitUIC") <> "12345" Or values("ReportingSeniorSSN") <> "000000009" Or values("Trait33") <> "3.0" Then Err.Raise 5, , "Extracted synthetic values differ."
    If SourceFileHash(path) <> before Then Err.Raise 5, , "Source file changed."
    PDFTargetProbe = PDFExtractionStatus(path) & "; " & values.Count & " mapped synthetic fields verified"
    Call PDFClearPreview: Exit Function
Failed:
    PDFTargetProbe = "FAIL: " & Err.Number & " / " & Err.Description
    Call PDFClearPreview
End Function
