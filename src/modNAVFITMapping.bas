Attribute VB_Name = "modNAVFITMapping"
Option Explicit
Public Function NAVFITMapRows(ByVal kind As String) As Collection
If kind <> "EVAL" And kind <> "CHIEFEVAL" And kind <> "FITREP" Then Err.Raise vbObjectError + 5100, "modNAVFITMapping.NAVFITMapRows", "Unsupported NAVFIT report type."
Set NAVFITMapRows = CatalogueNAVFITRows(kind)
End Function
