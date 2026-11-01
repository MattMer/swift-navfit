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
 Set f = Forms!frmUXSeniorEditor!uxRSCATable.Form: Set initialRows = CreateObject("Scripting.Dictionary")
 Set db = CurrentDb
 For Each grade In RSTableGrades()
  row = Array(Null, Null, Null, Null)
  Set r = db.OpenRecordset("SELECT TOP 1 BaselineID,CurrentRSCA,TraitCount,AsOfDate,TraitSum,SourceNote FROM RSCABaseline WHERE SeniorID=" & Nz(Forms!frmUXSeniorEditor!txtSeniorID, 0) & " AND Paygrade=" & QuoteSQL(CStr(grade)) & " ORDER BY AsOfDate DESC,BaselineID DESC", dbOpenSnapshot)
  If Not r.EOF Then
   row = Array(r!CurrentRSCA.Value, r!TraitCount.Value, r!AsOfDate.Value, r!TraitSum.Value)
   f.Controls("date" & grade).ControlTipText = Left$(Nz(r!SourceNote, "Historical observation") & IIf(IsNull(r!TraitSum), "; estimated from average and trait count", "; exact historical sum"), 255)
   f.Controls("src" & grade).Caption = RSBaselineSource(r!BaselineID)
  Else
   f.Controls("date" & grade).ControlTipText = "No baseline recorded for this paygrade."
   f.Controls("src" & grade).Caption = ""
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
 RSHistoryView Not Forms!frmUXSeniorEditor!lstBaselines.Visible
 RSViewTableHistory = True
End Function
Private Sub RSHistoryView(ByVal history As Boolean)
 Dim name As Variant
 With Forms!frmUXSeniorEditor
  !uxRSCATable.Visible = Not history: !cmdSaveRSTable.Visible = Not history
  For Each name In Array("lstBaselines", "cboHistoryGrade", "lblTrailSummary", "cmdOpenGroup"): .Controls(name).Visible = history: Next
  !cmdRSHistory.Caption = IIf(history, "Back to paygrade table", "View RSCA history")
 End With
 If history Then RSShowHistory
End Sub
'Where a paygrade's latest RSCA comes from: the entered BOL baseline, or that
'baseline plus the summary group completions posted on top of it.
Public Function RSBaselineSource(ByVal baselineID As Long) As String
 Dim id As Long, parent As Long, steps As Long
 id = baselineID
 Do
  parent = Nz(DLookup("ParentBaselineID", "RSCABaseline", "BaselineID=" & id), 0)
  If parent = 0 Or parent >= id Then Exit Do
  steps = steps + 1: id = parent
 Loop
 RSBaselineSource = "BOL " & Format$(DLookup("AsOfDate", "RSCABaseline", "BaselineID=" & id), "yyyy-mm-dd") & IIf(steps = 0, "", " + " & steps & " group" & IIf(steps = 1, "", "s"))
End Function
Private Function HistoryCell(ByVal text As String) As String
 HistoryCell = """" & Replace(Replace(Replace(text, """", """"""), vbCr, " "), vbLf, " ") & """"
End Function
Private Function HistoryRSCA(ByVal v As Variant) As String
 If Not IsNull(v(3)) And v(2) > 0 Then
  HistoryRSCA = Format$(v(3) / v(2), "0.00")
 ElseIf Not IsNull(v(4)) Then
  HistoryRSCA = Format$(v(4), "0.00") & IIf(v(2) > 0, " est.", "")
 End If
End Function
'One paygrade's RSCA history as a value list, oldest first. Each local entry names
'the summary group whose completion posted it. The current chain runs from the
'latest baseline back through ParentBaselineID to the last entered BOL baseline;
'entries outside it are earlier history.
Public Function RSHistoryList(ByVal seniorID As Long, ByVal grade As String, ByRef summary As String) As String
 Dim db As DAO.Database, r As DAO.Recordset, c As DAO.Recordset, info As Object, chain As Object, key As Variant, v As Variant
 Dim id As Long, latest As Long, root As Long, groups As Long, entry As String, added As String, reports As String, groupID As Long, text As String
 Set db = CurrentDb: Set info = CreateObject("Scripting.Dictionary"): Set chain = CreateObject("Scripting.Dictionary")
 text = "GroupID;Date;Entry;Reports;Traits added;Total traits;RSCA after"
 summary = "No RSCA history recorded for " & IIf(Len(grade) = 0, "this senior", grade) & "."
 If seniorID = 0 Or Len(grade) = 0 Then RSHistoryList = text: Exit Function
 Set r = db.OpenRecordset("SELECT BaselineID,ParentBaselineID,AsOfDate,TraitCount,TraitSum,CurrentRSCA,SourceNote FROM RSCABaseline WHERE SeniorID=" & seniorID & " AND Paygrade=" & QuoteSQL(grade) & " ORDER BY AsOfDate,BaselineID", dbOpenSnapshot)
 Do While Not r.EOF
  info.Add CLng(r!BaselineID), Array(CLng(Nz(r!ParentBaselineID, 0)), r!AsOfDate.Value, CLng(r!TraitCount), r!TraitSum.Value, r!CurrentRSCA.Value, Nz(r!SourceNote, ""))
  r.MoveNext
 Loop
 r.Close
 If info.Count = 0 Then RSHistoryList = text: Exit Function
 latest = RSLatestBaseline(seniorID, grade): id = latest
 Do While id > 0 And info.Exists(id) And Not chain.Exists(id)
  chain.Add id, True: root = id: id = info(id)(0)
 Loop
 For Each key In info.Keys
  v = info(key): groupID = 0: added = "": reports = ""
  If v(0) = 0 Then
   entry = IIf(Left$(v(5), 16) = "Swift local RSCA", "Local completion", "BOL baseline (entered)")
  Else
   Set c = db.OpenRecordset("SELECT TOP 1 c.CompletionID,c.GroupID,c.PriorCompletionID,g.Title FROM GroupCompletion AS c LEFT JOIN SummaryGroup AS g ON c.GroupID=g.GroupID WHERE c.PostedBaselineID=" & key & " ORDER BY c.CompletionID", dbOpenSnapshot)
   If c.EOF Then
    entry = "Local completion"
   Else
    groupID = c!groupID
    entry = IIf(IsNull(c!PriorCompletionID), "Group completed: ", "Group re-completed (correction): ") & Nz(c!Title, "group " & c!groupID)
    reports = CStr(DCount("*", "GroupCompletionReport", "CompletionID=" & c!CompletionID & " AND TraitCount>0"))
   End If
   c.Close
   If info.Exists(v(0)) Then added = CStr(v(2) - info(v(0))(2))
   If chain.Exists(key) Then groups = groups + 1
  End If
  If Not chain.Exists(key) Then entry = "Earlier: " & entry
  text = text & ";" & groupID & ";" & HistoryCell(Format$(v(1), "yyyy-mm-dd")) & ";" & HistoryCell(entry) & ";" & HistoryCell(reports) & ";" & HistoryCell(added) & ";" & v(2) & ";" & HistoryCell(HistoryRSCA(v))
 Next
 v = info(root)
 summary = grade & " BOL baseline " & Format$(v(1), "yyyy-mm-dd") & ": RSCA " & IIf(Len(HistoryRSCA(v)) > 0, HistoryRSCA(v), "none") & " (" & v(2) & " traits)"
 If latest <> root Then
  v = info(latest)
  summary = summary & "; now " & HistoryRSCA(v) & " (" & v(2) & " traits) after " & groups & " group completion" & IIf(groups = 1, "", "s") & "."
 Else
  summary = summary & ". No summary groups completed since."
 End If
 RSHistoryList = text
End Function
Public Function RSShowHistory() As Boolean
 Dim f As Access.Form, senior As Long, grade As String, grades As String, g As Variant, summary As String
 Set f = Forms!frmUXSeniorEditor: senior = Nz(f!txtSeniorID, 0)
 For Each g In RSTableGrades()
  If DCount("*", "RSCABaseline", "SeniorID=" & senior & " AND Paygrade=" & QuoteSQL(CStr(g))) > 0 Then grades = grades & IIf(Len(grades) > 0, ";", "") & g
 Next
 f!cboHistoryGrade.RowSource = grades
 grade = Nz(f!cboHistoryGrade, "")
 If Len(grade) = 0 Or InStr(";" & grades & ";", ";" & grade & ";") = 0 Then
  grade = Nz(MajorScalar("SELECT TOP 1 Paygrade FROM RSCABaseline WHERE SeniorID=" & senior & " ORDER BY AsOfDate DESC,BaselineID DESC"), "")
  If Len(grade) = 0 Then f!cboHistoryGrade = Null Else f!cboHistoryGrade = grade
 End If
 f!lstBaselines.RowSource = RSHistoryList(senior, grade, summary)
 f!lblTrailSummary.Caption = summary
 RSShowHistory = True
End Function
Public Function RSHistoryOpenGroup() As Boolean
 Dim id As Long
 id = Val(Nz(Forms!frmUXSeniorEditor!lstBaselines, 0))
 If id = 0 Then Forms!frmUXSeniorEditor!lblStatus.Caption = "Choose a group completion row first.": Exit Function
 If DCount("*", "SummaryGroup", "GroupID=" & id) = 0 Then Forms!frmUXSeniorEditor!lblStatus.Caption = "That summary group no longer exists.": Exit Function
 If Not RSLeaveSeniorEditor() Then Exit Function
 RSHistoryOpenGroup = UXShowGroupMembers(id)
End Function
Public Function RSTableResize() As Boolean
 Dim f As Access.Form, c As Access.Control, grade As Variant, width As Long, ratio As Double
 On Error GoTo Finished
 Set f = Forms!frmUXSeniorEditor!uxRSCATable.Form: width = f.InsideWidth - 350
 If width < 3000 Or f.Tag = CStr(width) Then Exit Function
 'Shrink to fit a narrow window; never stretch the numeric columns past their design.
 ratio = width / 10000: If ratio > 1 Then ratio = 1
 For Each grade In RSTableGrades()
  f.Controls("grade" & grade).Left = 100 * ratio: f.Controls("grade" & grade).Width = 900 * ratio
  f.Controls("avg" & grade).Left = 1100 * ratio: f.Controls("avg" & grade).Width = 1200 * ratio
  f.Controls("count" & grade).Left = 2400 * ratio: f.Controls("count" & grade).Width = 2100 * ratio
  f.Controls("date" & grade).Left = 4600 * ratio: f.Controls("date" & grade).Width = 1800 * ratio
  f.Controls("src" & grade).Left = 6500 * ratio: f.Controls("src" & grade).Width = 3400 * ratio
 Next
 f!headgrade.Left = 100 * ratio: f!headgrade.Width = 900 * ratio
 f!headavg.Left = 1100 * ratio: f!headavg.Width = 1200 * ratio
 f!headcount.Left = 2400 * ratio: f!headcount.Width = 2100 * ratio
 f!headdate.Left = 4600 * ratio: f!headdate.Width = 1800 * ratio
 f!headsrc.Left = 6500 * ratio: f!headsrc.Width = 3400 * ratio
 f.Tag = CStr(width): f.Width = width: RSTableResize = True
Finished:
End Function
Public Sub RSTablePage(ByVal page As Long)
 Dim name As Variant
 With Forms!frmUXSeniorEditor
  !uxRSCATable.Visible = (page = 3)
  For Each name In Array("lstBaselines", "cboHistoryGrade", "lblTrailSummary", "cmdOpenGroup"): .Controls(name).Visible = False: Next
  If page = 3 And Not shown Then RSTableLoad: shown = True
  !cmdRSHistory.Caption = "View RSCA history"
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
  For Each c In Array("grade", "avg", "count", "date", "src")
   f.Controls(c & grade).Top = 500 + i * 390
  Next
  f.Controls("avg" & grade).TabIndex = i * 3
  f.Controls("count" & grade).TabIndex = i * 3 + 1
  f.Controls("date" & grade).TabIndex = i * 3 + 2
 Next
 RSTableSort = True
End Function
