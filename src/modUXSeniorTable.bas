Attribute VB_Name = "modUXSeniorTable"
Option Explicit
Private rsSortKey As String, rsSortDescending As Boolean

Private initialRows As Object
Private shown As Boolean
Public Sub RSTableReset()
 shown = False: Set initialRows = Nothing
End Sub
Public Sub RSTableReloadVisible()
 If shown Then RSTableLoad
End Sub
Public Function RSTableGrades() As Variant
 RSTableGrades = Array("E1", "E2", "E3", "E4", "E5", "E6", "E7", "E8", "E9", "W1", "W2", "W3", "W4", "W5", "O1", "O2", "O3", "O4", "O5", "O6", "O7", "O8", "O9", "O10")
End Function
Private Function TableValue(ByVal name As String) As Variant
 Dim c As Access.Control, result As Variant
 Set c = Forms!frmUXSeniorEditor!uxRSCATable.Form.Controls(name)
 On Error GoTo Stored
 If Screen.ActiveControl Is c Then
  result = c.Text: GoTo Normalize
 End If
Stored:
 result = c.Value
Normalize:
 If Len(Nz(result, "")) = 0 Then
  result = Null
 ElseIf Left$(name, 3) = "avg" Or Left$(name, 5) = "count" Then
  If IsNumeric(result) Then result = CDbl(result)
 End If
 TableValue = result
End Function
' Load latest dated observations; prior baseline records are never overwritten.
Public Function RSTableLoad() As Boolean
 Dim grade As Variant, r As DAO.Recordset, db As DAO.Database, f As Access.Form, row As Variant
 If Not UXHasControl(Forms!frmUXSeniorEditor, "uxRSCATable") Then Exit Function
 Set f = Forms!frmUXSeniorEditor!uxRSCATable.Form: Set initialRows = CreateObject("Scripting.Dictionary")
 Set db = CurrentDb
 For Each grade In RSTableGrades()
  row = Array(Null, Null, Null, Null)
  Set r = db.OpenRecordset("SELECT TOP 1 CurrentRSCA,TraitCount,AsOfDate,TraitSum,SourceNote FROM RSCABaseline WHERE SeniorID=" & Nz(Forms!frmUXSeniorEditor!txtSeniorID, 0) & " AND Paygrade=" & QuoteSQL(CStr(grade)) & " ORDER BY AsOfDate DESC,BaselineID DESC", dbOpenSnapshot)
  If Not r.EOF Then
   row = Array(r!CurrentRSCA.Value, r!TraitCount.Value, r!AsOfDate.Value, r!TraitSum.Value)
   f.Controls("date" & grade).ControlTipText = Left$(Nz(r!SourceNote, "Historical observation") & IIf(IsNull(r!TraitSum), "; estimated from average and trait count", "; exact historical sum"), 255)
  Else
   f.Controls("date" & grade).ControlTipText = "No baseline recorded for this paygrade."
  End If
  r.Close
  f.Controls("avg" & grade) = row(0): f.Controls("count" & grade) = row(1): f.Controls("date" & grade) = row(2)
  initialRows.Add CStr(grade), row
 Next
 RSTableLoad = True
End Function
Public Function RSTableDirty() As Boolean
 Dim grade As Variant, row As Variant
 If initialRows Is Nothing Then Exit Function
 For Each grade In RSTableGrades()
  row = initialRows(CStr(grade))
  If Not SameFieldValue(TableValue("avg" & grade), row(0)) Or Not SameFieldValue(TableValue("count" & grade), row(1)) Or Not SameFieldValue(UXDateEntry(TableValue("date" & grade)), UXDateEntry(row(2))) Then RSTableDirty = True: Exit Function
 Next
End Function
' Validate all edited rows before appending any; roll back the whole table on failure.
Public Function RSSaveTable() As Boolean
 Dim grade As Variant, row As Variant, average As Variant, count As Variant, dateText As String, verified As Date, total As Variant
 Dim plans As New Collection, item As Variant, w As DAO.Workspace, trans As Boolean, id As Long, message As String
 On Error GoTo Failed
 If Nz(Forms!frmUXSeniorEditor!txtSeniorID, 0) = 0 Or RSSeniorDirty() Then Err.Raise vbObjectError + 8300, "modUXSeniorTable.RSSaveTable", "Save the reporting senior record details first."
 For Each grade In RSTableGrades()
  row = initialRows(CStr(grade)): average = TableValue("avg" & grade): count = TableValue("count" & grade): dateText = Nz(UXDateEntry(TableValue("date" & grade)), "")
  If Not SameFieldValue(average, row(0)) Or Not SameFieldValue(count, row(1)) Or Not SameFieldValue(dateText, Nz(UXDateEntry(row(2)), "")) Then
   If Len(Nz(count, "")) = 0 Or Not IsNumeric(count) Then Err.Raise vbObjectError + 8301, "modUXSeniorTable.RSSaveTable", grade & ": enter the previous graded trait count."
   If CDbl(count) < 0 Or CDbl(count) > 2147483647# Or CDbl(count) <> Fix(CDbl(count)) Then Err.Raise vbObjectError + 8301, "modUXSeniorTable.RSSaveTable", grade & ": count must be a nonnegative whole number."
   If Len(Nz(average, "")) = 0 Then average = Null
   If Not dateText Like "####-##-##" Then Err.Raise vbObjectError + 8302, "modUXSeniorTable.RSSaveTable", grade & ": enter the date verified/pulled from BOL."
   verified = DateSerial(CInt(Left$(dateText, 4)), CInt(Mid$(dateText, 6, 2)), CInt(Right$(dateText, 2)))
   If Format$(verified, "yyyy-mm-dd") <> dateText Then Err.Raise vbObjectError + 8302, "modUXSeniorTable.RSSaveTable", grade & ": enter a valid date."
   total = Null
   If SameFieldValue(average, row(0)) And SameFieldValue(count, row(1)) Then total = row(3)
   RSValidateBaseline CStr(grade), average, CLng(count), total
   plans.Add Array(CStr(grade), verified, average, CLng(count), total)
  End If
 Next
 If plans.Count = 0 Then RSSaveTable = True: Exit Function
 Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
 For Each item In plans
  id = RSAddBaseline(CLng(Forms!frmUXSeniorEditor!txtSeniorID), CStr(item(0)), CDate(item(1)), item(2), CLng(item(3)), item(4), "Manually verified/pulled from BOL")
 Next
 w.CommitTrans: trans = False
 RSReloadBaselines
 Forms!frmUXSeniorEditor!lblStatus.Caption = "Saved " & plans.Count & " paygrade rows. Earlier baselines remain in history. Counts are traits, not reports."
 RSSaveTable = True: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modUXSeniorTable", "RSSaveTable", "Execute", Err.Number, Err.Description, Err.Source, IIf(trans, "Active", "Not active"))
 message = CStr(swiftPrimary("Description"))
 On Error Resume Next
 If trans Then SwiftRollback w, swiftPrimary
 On Error GoTo 0
 Forms!frmUXSeniorEditor!lblStatus.Caption = SwiftFailureText(swiftPrimary) & " Table entries retained."
End Function
Public Function RSSaveTableButton() As Boolean
 Forms!frmUXSeniorEditor!cmdSave.SetFocus
 RSSaveTableButton = RSSaveTable()
End Function
Public Function RSViewTableHistory() As Boolean
 Forms!frmUXSeniorEditor!cmdSave.SetFocus
 With Forms!frmUXSeniorEditor
  !uxRSCATable.Visible = Not !uxRSCATable.Visible
  !lstBaselines.Visible = Not !uxRSCATable.Visible
  !cmdRSHistory.Caption = IIf(!uxRSCATable.Visible, "View prior baselines", "Back to paygrade table")
 End With
 RSViewTableHistory = True
End Function
Public Function RSTableResize() As Boolean
 Dim f As Access.Form, c As Access.Control, grade As Variant, width As Long, ratio As Double
 On Error GoTo Finished
 Set f = Forms!frmUXSeniorEditor!uxRSCATable.Form: width = f.InsideWidth - 350
 If width < 3000 Or f.Tag = CStr(width) Then Exit Function
 'Shrink to fit a narrow window; never stretch the numeric columns past their design.
 ratio = width / 10000: If ratio > 1 Then ratio = 1
 For Each grade In RSTableGrades()
  f.Controls("grade" & grade).Left = 100 * ratio: f.Controls("grade" & grade).Width = 800 * ratio
  f.Controls("avg" & grade).Left = 1000 * ratio: f.Controls("avg" & grade).Width = 1600 * ratio
  f.Controls("count" & grade).Left = 2900 * ratio: f.Controls("count" & grade).Width = 2600 * ratio
  f.Controls("date" & grade).Left = 5800 * ratio: f.Controls("date" & grade).Width = 3700 * ratio
 Next
 f!headgrade.Left = 100 * ratio: f!headgrade.Width = 800 * ratio
 f!headavg.Left = 1000 * ratio: f!headavg.Width = 1600 * ratio
 f!headcount.Left = 2900 * ratio: f!headcount.Width = 2600 * ratio
 f!headdate.Left = 5800 * ratio: f!headdate.Width = 3700 * ratio
 f.Tag = CStr(width): f.Width = width: RSTableResize = True
Finished:
End Function
Public Sub RSTablePage(ByVal page As Long)
 With Forms!frmUXSeniorEditor
  If Not UXHasControl(Forms!frmUXSeniorEditor, "uxRSCATable") Then Exit Sub
  !uxRSCATable.Visible = (page = 3): !lstBaselines.Visible = False
  If page = 3 And Not shown Then RSTableLoad: shown = True
  !cmdRSHistory.Caption = "View prior baselines"
 End With
End Sub
Public Function RSTableChanged() As Boolean
 Forms!frmUXSeniorEditor!lblStatus.Caption = "Unsaved RSCA table entries. Save RSCA table to record verified baselines."
 RSTableChanged = True
End Function

Public Function RSTableSort(ByVal key As String) As Boolean
 Dim grades As Variant, order() As Long, values() As Variant, i As Long, j As Long, temp As Long, f As Access.Form, grade As String, c As Variant
 If key <> "grade" And key <> "avg" And key <> "count" And key <> "date" Then Exit Function
 Set f = Forms!frmUXSeniorEditor!uxRSCATable.Form: grades = RSTableGrades()
 ReDim order(0 To UBound(grades)): ReDim values(0 To UBound(grades))
 If rsSortKey = key Then rsSortDescending = Not rsSortDescending Else rsSortDescending = False
 rsSortKey = key
 For i = 0 To UBound(grades)
  order(i) = i
  If key = "grade" Then values(i) = i Else values(i) = TableValue(key & grades(i))
 Next
 For i = 1 To UBound(order)
  temp = order(i): j = i - 1
  Do While j >= 0
   If UXCompareCells(values(order(j)), values(temp)) * IIf(rsSortDescending, -1, 1) <= 0 Then Exit Do
   order(j + 1) = order(j): j = j - 1
  Loop
  order(j + 1) = temp
 Next
 For i = 0 To UBound(order)
  grade = grades(order(i))
  For Each c In Array("grade", "avg", "count", "date")
   f.Controls(c & grade).Top = 500 + i * 390
  Next
  f.Controls("avg" & grade).TabIndex = i * 3
  f.Controls("count" & grade).TabIndex = i * 3 + 1
  f.Controls("date" & grade).TabIndex = i * 3 + 2
 Next
 RSTableSort = True
End Function
