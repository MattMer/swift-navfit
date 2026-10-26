Attribute VB_Name = "modTableSort"
Option Explicit
Private states As Object
Private headersBusy As Boolean
Private headings As Object
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetScrollPos Lib "user32" (ByVal hwnd As LongPtr, ByVal bar As Long) As Long
Private Declare PtrSafe Function GetDpiForWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Function SortForm(ByVal name As String) As Access.Form
 If name = "frmUXEditorFields" Then
  Set SortForm = Forms!frmEditor!uxFieldSurface.Form
 ElseIf name = "frmReportTools" Then
  Set SortForm = UXToolsForm()
 ElseIf name = "frmReportReviewPanel" Then
  Set SortForm = UXReviewForm()
 Else
  Set SortForm = Forms(name)
 End If
End Function
Public Function UXCompareCells(ByVal a As Variant, ByVal b As Variant) As Long
 If Len(Nz(a, "")) = 0 Then
  If Len(Nz(b, "")) > 0 Then UXCompareCells = -1
 ElseIf Len(Nz(b, "")) = 0 Then
  UXCompareCells = 1
 ElseIf IsNumeric(a) And IsNumeric(b) Then
  If CDbl(a) < CDbl(b) Then UXCompareCells = -1 Else If CDbl(a) > CDbl(b) Then UXCompareCells = 1
 ElseIf IsDate(a) And IsDate(b) Then
  If CDate(a) < CDate(b) Then UXCompareCells = -1 Else If CDate(a) > CDate(b) Then UXCompareCells = 1
 Else
  UXCompareCells = StrComp(CStr(a), CStr(b), vbTextCompare)
 End If
End Function
Private Function ListTokens(ByVal source As String) As Collection
 Dim result As New Collection, i As Long, quoted As Boolean, token As String, ch As String
 For i = 1 To Len(source)
  ch = Mid$(source, i, 1)
  If ch = """" Then
   If quoted And Mid$(source, i + 1, 1) = """" Then token = token & ch: i = i + 1 Else quoted = Not quoted
  ElseIf ch = ";" And Not quoted Then
   result.Add token: token = ""
  Else
   token = token & ch
  End If
 Next
 If Len(source) > 0 Then result.Add token
 Set ListTokens = result
End Function
Private Function CellQuote(ByVal value As Variant) As String
 CellQuote = """" & Replace(Nz(value, ""), """", """""") & """"
End Function
Private Sub SortValueList(ByVal c As Access.Control, ByVal column As Long, ByVal direction As Long)
 Dim tokens As Collection, rows() As Variant, count As Long, cols As Long, start As Long, i As Long, j As Long, k As Long, temp As Variant, text As String
 Set tokens = ListTokens(c.RowSource): cols = c.ColumnCount
 If tokens.Count = 0 Then Exit Sub
 If tokens.Count Mod cols <> 0 Then Err.Raise 5, "modTableSort.SortValueList", "List cells are incomplete; refresh the table before sorting."
 count = tokens.Count \ cols: ReDim rows(0 To count - 1)
 For i = 0 To count - 1
  ReDim temp(0 To cols - 1)
  For k = 0 To cols - 1: temp(k) = tokens(i * cols + k + 1): Next
  rows(i) = temp
 Next
 If c.ColumnHeads Then start = 1
 'Stable insertion sort keeps equal-value rows in their prior displayed order.
 For i = start + 1 To count - 1
  temp = rows(i): j = i - 1
  Do While j >= start
   If UXCompareCells(rows(j)(column), temp(column)) * direction <= 0 Then Exit Do
   rows(j + 1) = rows(j): j = j - 1
  Loop
  rows(j + 1) = temp
 Next
 For i = 0 To count - 1
  For k = 0 To cols - 1
   If Len(text) > 0 Then text = text & ";"
   text = text & CellQuote(rows(i)(k))
  Next
 Next
 c.RowSource = text
End Sub
Public Function UXSortList(ByVal formName As String, ByVal controlName As String, ByVal column As Long) As Boolean
 Dim c As Access.Control, db As DAO.Database, q As DAO.QueryDef, state As Object, token As String, source As String, name As String, key As String, sql As String, direction As Long, chosen As Object, i As Long, bound As Long, selected As Variant
 On Error GoTo Failed
 Set c = SortForm(formName).Controls(controlName)
 If column < 0 Or column >= c.ColumnCount Then Exit Function
 If states Is Nothing Then Set states = CreateObject("Scripting.Dictionary")
 token = formName & ":" & controlName: direction = 1
 If states.Exists(token) Then
  Set state = states(token)
  If state("Column") = column And c.RowSource = state("Last") Then direction = -CLng(state("Direction"))
 Else
  Set state = CreateObject("Scripting.Dictionary"): state.Add "Column", -1: state.Add "Direction", 1: state.Add "Base", "": state.Add "Last", "": states.Add token, state
 End If
 Set chosen = CreateObject("Scripting.Dictionary"): bound = c.BoundColumn - 1
 If bound < 0 Then bound = 0
 If c.MultiSelect > 0 Then
  For i = 0 To c.ListCount - 1: If c.Selected(i) Then chosen(CStr(Nz(c.Column(bound, i), ""))) = True
  Next
 Else
  selected = c.Value
 End If
 If c.RowSourceType = "Value List" Then
  SortValueList c, column, direction
 Else
  Set db = CurrentDb: source = c.RowSource
  If source = "" Then Exit Function
  name = Left$("uxSort_" & formName & "_" & controlName, 64)
  If source <> state("Last") Or state("Base") = "" Then
   If InStr(1, source, "SELECT", vbTextCompare) = 0 Then source = "SELECT * FROM [" & Replace(source, "]", "]]") & "]"
   On Error Resume Next: Set q = db.QueryDefs(name): On Error GoTo Failed
   If q Is Nothing Then Set q = db.CreateQueryDef(name, source) Else q.SQL = source
   state("Base") = source
  Else
   Set q = db.QueryDefs(name)
  End If
  key = q.Fields(column).Name
  sql = "SELECT * FROM [" & name & "] ORDER BY [" & Replace(key, "]", "]]") & "] " & IIf(direction = 1, "ASC", "DESC")
  state("Last") = sql: c.RowSource = sql: q.Close
 End If
 c.Requery
 'Changing column widths can clear Access multi-selection; lay out headings first.
 UXSortHeaders SortForm(formName)
 If c.MultiSelect > 0 Then
  For i = 0 To c.ListCount - 1: c.Selected(i) = chosen.Exists(CStr(Nz(c.Column(bound, i), ""))): Next
 ElseIf Not IsNull(selected) Then
  c.Value = selected
 End If
 state("Column") = column: state("Direction") = direction: state("Last") = c.RowSource
 c.ControlTipText = "Sorted column " & (column + 1) & IIf(direction = 1, " ascending", " descending") & ". Click its header to reverse."
 UXSortList = True: Exit Function
Failed:
 If Not c Is Nothing Then c.ControlTipText = "Unable to sort: " & Err.Description
End Function
Public Function UXSortNativeHeader(ByVal formName As String, ByVal controlName As String, ByVal x As Single, ByVal y As Single) As Boolean
 Dim c As Access.Control, widths As Variant, column As Long, position As Long, size As Long, hwnd As LongPtr, dpi As Long
 Set c = SortForm(formName).Controls(controlName)
 If y > c.FontSize * 20 + 90 Or y < 0 Then Exit Function
 position = CLng(x)
 On Error Resume Next
 hwnd = GetFocus(): dpi = GetDpiForWindow(hwnd)
 If dpi > 0 Then position = position + CLng(GetScrollPos(hwnd, 0) * 1440# / dpi)
 On Error GoTo 0
 widths = Split(c.ColumnWidths, ";")
 For column = 0 To c.ColumnCount - 1
  size = c.Width \ c.ColumnCount
  If column <= UBound(widths) Then
   If Len(widths(column)) > 0 Then size = Val(widths(column))
  End If
  If size > 0 Then
   If position < size Then UXSortNativeHeader = UXSortList(formName, controlName, column): Exit Function
   position = position - size
  End If
 Next
End Function
Private Function Heading(ByVal f As Access.Form, ByVal c As Access.Control, ByVal column As Long) As String
 Dim labels As Variant, db As DAO.Database, q As DAO.QueryDef, source As String
 Select Case c.Name
  Case "lstImportReports": labels = Array("ID", "Report type", "Import state", "Member")
  Case "lstImportFields": labels = Array("ID", "Field", "Type", "Value state")
  Case "lstCompareFields": labels = Array("Key", "Report block", "Change")
  Case "lstFindings": labels = Array("Key", "Severity", "Finding")
  Case "lstHistory": labels = Array("ID", "Source", "Report type", "Period", "ITA", "Promotion")
  Case "lstCounseling"
   If f.Name = "frmMyWork" Then labels = Array("ID", "Member", "Rank/rate", "Status", "Due month", "Last counseled", "Note") Else labels = Array("ID", "Date", "Outcome", "Counselor or reason", "Worksheet")
  Case "lstChoices": labels = Array("Field", "Change", "Decision")
  Case "lstCatalog": labels = Array("ID", "Rate", "Type", "Example")
  Case "lstRows": labels = Array("ID", "Member", "Designator", "RUIC", "Action")
  Case "lstGroups"
   If f.Name = "frmUXNativeGroups" Then labels = Array("ID", "Folder / group", "Report type", "Paygrade", "Reports")
 End Select
 If IsArray(labels) Then
  If column <= UBound(labels) Then Heading = labels(column): Exit Function
 End If
 On Error GoTo Fallback
 If c.RowSourceType <> "Value List" And Len(c.RowSource) > 0 Then
  source = c.RowSource: If InStr(1, source, "SELECT", vbTextCompare) = 0 Then source = "SELECT * FROM [" & source & "]"
  Set db = CurrentDb: Set q = db.CreateQueryDef("", source): Heading = q.Fields(column).Name: q.Close: Exit Function
 End If
Fallback:
 Heading = "Column " & (column + 1)
End Function
'A SQL list's heading needs a temporary QueryDef; reuse it until the row source changes.
Private Function CachedHeading(ByVal f As Access.Form, ByVal c As Access.Control, ByVal column As Long) As String
 Dim key As String
 If c.RowSourceType = "Value List" Then CachedHeading = Heading(f, c, column): Exit Function
 If headings Is Nothing Then Set headings = CreateObject("Scripting.Dictionary")
 key = f.Name & "|" & c.Name & "|" & column & "|" & c.RowSource
 If Not headings.Exists(key) Then headings.Add key, Heading(f, c, column)
 CachedHeading = headings(key)
End Function
Private Function SameWidths(ByVal current As String, ByVal wanted As String) As Boolean
 Dim a As Variant, b As Variant, i As Long
 a = Split(current, ";"): b = Split(wanted, ";")
 If UBound(a) <> UBound(b) Then Exit Function
 For i = 0 To UBound(a)
  If Abs(Val(a(i)) - Val(b(i))) > 2 Then Exit Function
 Next
 SameWidths = True
End Function
'Custom headings for lists that did not expose native headers. Their original
'Value List data stays intact; no first data row is consumed as a heading.
Public Function UXSortHeaders(ByVal f As Access.Form) As Boolean
 Dim h As Access.Control, c As Access.Control, parts As Variant, widths As Variant, i As Long, column As Long, x As Long, total As Double, size As Long, sizes As String
 If headersBusy Then Exit Function
 headersBusy = True
 On Error GoTo Finished
 For Each c In f.Controls
  'VBA And is not short-circuiting: other controls have no ColumnHeads property.
  If c.ControlType = acListBox Then
   If Not c.ColumnHeads Then
   widths = Split(c.ColumnWidths, ";"): total = 0
   For i = 0 To c.ColumnCount - 1
    size = c.Width \ c.ColumnCount
    If i <= UBound(widths) Then
     If Len(widths(i)) > 0 Then size = Val(widths(i))
    End If
    total = total + size
   Next
   x = c.Left: sizes = ""
   For column = 0 To c.ColumnCount - 1
    size = c.Width \ c.ColumnCount
    If column <= UBound(widths) Then
     If Len(widths(column)) > 0 Then size = Val(widths(column))
    End If
    If total > 0 And size > 0 Then size = CLng(size * (c.Width - 300) / total)
    If column > 0 Then sizes = sizes & ";"
    sizes = sizes & size
    If UXHasControl(f, "sort_" & c.Name & "_" & column) Then
     Set h = f.Controls("sort_" & c.Name & "_" & column)
     UXVisualSet h, "Visible", (c.Visible And size > 0)
     If size > 0 Then
      UXVisualSet h, "Left", x: UXVisualSet h, "Top", c.Top - 330: UXVisualSet h, "Width", size: UXVisualSet h, "Height", 300
      UXVisualSet h, "Caption", CachedHeading(f, c, column)
      UXVisualSet h, "ControlTipText", "Click to sort ascending; click again for descending."
     End If
    End If
    x = x + size
   Next
   'Rewriting ColumnWidths redraws the list; rescaling may drift by a twip or two.
   If total > 0 And UXHasControl(f, "sort_" & c.Name & "_0") Then
    If Not SameWidths(c.ColumnWidths, sizes) Then c.ColumnWidths = sizes
   End If
   End If
  End If
 Next
 UXSortHeaders = True
Finished:
 headersBusy = False
End Function
Public Function UXSortActivate(ByVal formName As String) As Boolean
 Dim f As Access.Form, original As String
 Set f = SortForm(formName)
 If UXHasControl(f, "uxSortActivateOriginal") Then
  original = Nz(f!uxSortActivateOriginal, "")
  If Left$(original, 1) = "=" Then Application.Eval Mid$(original, 2)
 End If
 UXSortActivate = UXSortHeaders(f)
End Function
