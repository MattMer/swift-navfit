Attribute VB_Name = "modReportPageMap"
Option Explicit
Public Function ReportPageForField(ByVal kind As String, ByVal key As String) As Long
ReportPageForField = CataloguePage(kind, key)
End Function
Public Function ReportPageRows(ByVal kind As String, ByVal page As Long) As Collection
Set ReportPageRows = CataloguePageRows(kind, page)
End Function
