Attribute VB_Name = "modPDFMapping"
Option Explicit
Public Function PDFMapRows(ByVal kind As String) As Collection
If kind <> "EVAL" And kind <> "CHIEFEVAL" And kind <> "FITREP" Then Err.Raise vbObjectError + 4720, , "Unsupported PDF report type."
Set PDFMapRows = CataloguePDFRows(kind)
End Function
Public Function PDFMappingCount(ByVal kind As String) As Long
PDFMappingCount = PDFMapRows(kind).Count
End Function
