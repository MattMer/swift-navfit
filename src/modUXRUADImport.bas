Attribute VB_Name = "modUXRUADImport"
Option Explicit
Private loading As Boolean
Public Function UXRUADImportOpen() As Boolean
    DoCmd.OpenForm "frmUXRUADImport"
    UXRUADImportOpen = True
End Function
Public Function UXRUADImportChoose() As Boolean
    Dim dialog As Object, f As Access.Form, count As Long
    On Error GoTo Failed
    Set f = Forms!frmUXRUADImport: f!cmdApply.Enabled = False
    Set dialog = Application.FileDialog(3)
    dialog.Title = "Choose the command RUAD CSV": dialog.AllowMultiSelect = False
    dialog.Filters.Clear: dialog.Filters.Add "RUAD CSV", "*.csv"
    If dialog.Show <> -1 Then Exit Function
    f!lstRows.RowSource = "": f!txtDetails = Null
    count = RUADLoad(CStr(dialog.SelectedItems(1)))
    f!txtPath = CStr(dialog.SelectedItems(1))
    f!lblStatus.Caption = "Choose actions for " & count & " rows. Vacant billets are skipped. Names suggest matches; they do not establish identity."
    UXRUADImportRows: UXRUADImportChoose = True: Exit Function
Failed:
    f!lblStatus.Caption = Err.Description
End Function
Public Function UXRUADImportRows() As Boolean
    Dim f As Access.Form, i As Long, row As Object, data As Object, label As String, db As DAO.Database, r As DAO.Recordset
    Set f = Forms!frmUXRUADImport: loading = True
    f!lstRows.RowSource = ""
    For i = 1 To RUADCount()
        Set row = RUADRow(i): Set data = row("Data")
        label = IIf(row("Vacant"), "Vacant billet", CStr(data("NAME")))
        f!lstRows.AddItem CStr(i) & ";" & RUADListText(label) & ";" & RUADListText(CStr(data("DESG"))) & ";" & RUADListText(CStr(data("RUIC"))) & ";" & RUADListText(CStr(row("Action")))
    Next
    loading = False: UXRUADImportRows = True
End Function
Private Function RUADListText(ByVal value As String) As String
    RUADListText = """" & Replace(Replace(Replace(value, """", """"""), vbCr, " "), vbLf, " ") & """"
End Function
Public Function UXRUADImportSelect() As Boolean
    Dim f As Access.Form, row As Object, data As Object
    If loading Then Exit Function
    Set f = Forms!frmUXRUADImport
    If IsNull(f!lstRows) Then Exit Function
    loading = True: Set row = RUADRow(CLng(f!lstRows)): Set data = row("Data")
    f!cboAction = row("Action"): f!cboMatch = IIf(row("MemberID") = 0, Null, row("MemberID")): f!chkReplaceManual = row("ReplaceManual")
    f!cboMatch.RowSource = "SELECT MemberID,FullName,RankRate FROM CommandMember WHERE IsArchived=False ORDER BY IIf(FullName=" & QuoteSQL(CStr(data("NAME"))) & ",0,1),FullName,MemberID"
    f!cboMatch.Requery
    f!txtDetails = row("Detail"): f!cboAction.Enabled = Not row("Vacant"): f!cboMatch.Enabled = Not row("Vacant"): f!chkReplaceManual.Enabled = Not row("Vacant")
    loading = False: UXRUADImportSelect = True
End Function
Public Function UXRUADImportSelection() As Boolean
    Dim f As Access.Form, selected As Long
    If loading Then Exit Function
    On Error GoTo Failed
    Set f = Forms!frmUXRUADImport: f!cmdApply.Enabled = False
    If IsNull(f!lstRows) Then Exit Function
    selected = CLng(f!lstRows)
    If Nz(f!cboAction, "Skip") = "Refresh existing" And Nz(f!cboMatch, 0) = 0 Then
        f!lblStatus.Caption = "Choose the existing member to match this row before previewing.": Exit Function
    End If
    If Nz(f!cboAction, "Skip") = "Skip" And Nz(f!cboMatch, 0) > 0 Then f!cboAction = "Refresh existing"
    RUADChoose selected, Nz(f!cboAction, "Skip"), Nz(f!cboMatch, 0), Nz(f!chkReplaceManual, False)
    UXRUADImportRows: f!lstRows = selected: UXRUADImportSelect
    f!lblStatus.Caption = "Selection changed. Preview again before Apply.": UXRUADImportSelection = True: Exit Function
Failed:
    f!lblStatus.Caption = Err.Description
End Function
Public Function UXRUADImportNew() As Boolean
    Dim i As Long, row As Object, data As Object, count As Long
    On Error GoTo Failed
    Forms!frmUXRUADImport!cmdApply.Enabled = False
    For i = 1 To RUADCount()
        Set row = RUADRow(i): Set data = row("Data")
        If Not row("Vacant") And row("Action") = "Skip" Then
            If DCount("*", "CommandMember", "FullName=" & QuoteSQL(CStr(data("NAME")))) = 0 Then
                RUADChoose i, "Create new", 0: count = count + 1
            End If
        End If
    Next
    UXRUADImportRows
    Forms!frmUXRUADImport!lblStatus.Caption = count & " new profiles selected. Existing names still require explicit matching. Preview before Apply."
    UXRUADImportNew = True: Exit Function
Failed:
    Forms!frmUXRUADImport!lblStatus.Caption = Err.Description
End Function
Public Function UXRUADImportPreview() As Boolean
    On Error GoTo Failed
    Forms!frmUXRUADImport!cmdApply.Enabled = False
    Forms!frmUXRUADImport!lblStatus.Caption = RUADPreview()
    If Not IsNull(Forms!frmUXRUADImport!lstRows) Then UXRUADImportSelect
    Forms!frmUXRUADImport!cmdApply.Enabled = True: UXRUADImportPreview = True: Exit Function
Failed:
    Forms!frmUXRUADImport!lblStatus.Caption = Err.Description
End Function
Public Function UXRUADImportApply() As Boolean
    Dim count As Long
    On Error GoTo Failed
    Forms!frmUXRUADImport!cmdApply.Enabled = False
    count = RUADApply()
    Forms!frmUXRUADImport!lblStatus.Caption = "Results: " & count & " profiles applied. Reports, SSNs, links and prior evaluations preserved. Operation recorded in History."
    If CurrentProject.AllForms("frmUXDirectory").IsLoaded Then UXDirectoryLoad
    UXRUADImportApply = True: Exit Function
Failed:
    Forms!frmUXRUADImport!lblStatus.Caption = Err.Description
End Function
