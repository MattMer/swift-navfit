Attribute VB_Name = "NativeDatesHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8390, , message
 n = n + 1
End Sub
Public Function TestNativeDates(ByVal source As String) As String
 Dim groups As Object, picks As New Collection, key As Variant, item As Object, nativeID As Variant, incoming As Object, hash As String, imported As Collection
 Dim db As DAO.Database, rows As DAO.Recordset, values As Object, field As Variant, raw As Variant, r As DAO.Recordset, original As DAO.Database
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC TIMESTAMP IMPORT", "Tester"
 Check NAVFITDecode(#10/16/2026 1:23:45 PM#, "date") = "2026-10-16", "Timed date rejected or shifted"
 Check NAVFITDecode(#10/16/2026 11:59:59 PM#, "date") = "2026-10-16", "Late time rolled calendar day"
 Check IsNull(NAVFITDecode(Null, "date")), "Null date changed"
 OpenReviewFile source, False
 Set groups = NAVFITReviewGroupChoices()
 Check groups.Count = 3, "Timed folder defaults hid native groups"
 For Each key In groups.Keys
  Set item = groups(key): picks.Add CStr(key): hash = item("SourceHash")
  Check item("Defaults")("PeriodStart") = "2026-01-01", "Folder date did not normalize"
  For Each nativeID In Split(item("Reports"), ",")
   Set incoming = FullNAVFITImportValues(CLng(nativeID))
   Check Format$(incoming("PeriodEnd"), "yyyy-mm-dd") = "2026-09-30", "Report date did not normalize"
  Next
 Next
 Set imported = ImportNAVFITGroups(picks, hash, 2000)
 Check imported.Count = 3 And DCount("*", "Report") = 3, "Timed reports did not import together"
 Set db = CurrentDb: Set original = DBEngine.OpenDatabase(source, False, True)
 Set rows = db.OpenRecordset("SELECT i.ImportReceiptID,i.ReportID,i.ExternalReportID FROM ImportReceipt AS i ORDER BY i.ImportReceiptID", dbOpenSnapshot)
 Do While Not rows.EOF
  Set values = SavedFields(rows!ReportID)
  Set r = original.OpenRecordset("SELECT * FROM Reports WHERE ReportID=" & rows!ExternalReportID, dbOpenSnapshot)
  For Each field In Array("FromDate|PeriodStart", "ToDate|PeriodEnd", "DateReported|DateReported", "PhysicalReadinessDt|ReadinessDate", "RaterDate|RaterDate", "SeniorRaterDate|SeniorRaterDate")
   Dim parts As Variant
   parts = Split(CStr(field), "|")
   If IsNull(r.Fields(parts(0))) Then
    Check IsNull(values(parts(1))), "Null date lost"
   Else
    Check Format$(values(parts(1)), "yyyy-mm-dd") = Format$(r.Fields(parts(0)), "yyyy-mm-dd"), "Saved day differs: " & parts(0)
    raw = MajorScalar("SELECT TOP 1 RawValue FROM ImportRawValue WHERE ImportReceiptID=" & rows!ImportReceiptID & " AND SourceTable='Reports' AND FieldKey=" & QuoteSQL(CStr(parts(0))) & " AND IsNullValue=False")
    Check Not IsNull(raw), "Raw timestamp absent"
    Check Abs(CDbl(raw) - CDbl(r.Fields(parts(0)))) < 0.00000001, "Raw timestamp changed"
   End If
  Next
  r.Close: rows.MoveNext
 Loop
 rows.Close: original.Close: CloseReviewFile
 TestNativeDates = "PASS: " & n & " timestamp import, date-only values and raw provenance assertions": Exit Function
Failed:
 TestNativeDates = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
