Attribute VB_Name = "modUXNavyWriter"
Option Explicit
Public Function UXLoadNavyWriter() As Boolean
    Dim i As Long, term As String, title As String, body As String, rate As Variant, f As Access.Form
    Set f = Forms!frmUXNavyWriter
    If Len(f!cboRate.RowSource) = 0 Then
        f!cboRate.RowSourceType = "Value List": f!cboRate.RowSource = "ALL;All rates;GENERAL;General / any rate"
        For Each rate In NavyWriterRates(): f!cboRate.AddItem CStr(rate) & ";" & CStr(rate): Next
        f!cboRate = "ALL"
    End If
    term = Nz(Forms!frmUXNavyWriter!txtSearch, "")
    f!lstCatalog = Null: f!txtBody = "": f!txtSource = "": f!cmdImport.Enabled = False: f!cmdSource.Enabled = False
    Forms!frmUXNavyWriter!lstCatalog.RowSourceType = "Value List": Forms!frmUXNavyWriter!lstCatalog.RowSource = ""
    For i = 0 To NavyWriterCount() - 1
        title = NavyWriterValue(i, 0): body = NavyWriterValue(i, 1)
        If NavyWriterMatches(i, Nz(f!cboRate, "ALL"), term) Then f!lstCatalog.AddItem CStr(i) & ";" & Replace(NavyWriterValue(i, 3), "|", ", ") & ";" & NavyWriterValue(i, 4) & ";" & Replace(title, ";", ",")
    Next
    f!lblStatus.Caption = f!lstCatalog.ListCount & " entries. Select an excerpt to adapt and add, or a collection to open its full source."
    UXLoadNavyWriter = True
End Function
Public Function NavyWriterMatches(ByVal row As Long, ByVal rate As String, ByVal term As String) As Boolean
    If row < 0 Or row >= NavyWriterCount() Then Exit Function
    If rate <> "ALL" And InStr(1, "|" & NavyWriterValue(row, 3) & "|", "|" & rate & "|", vbBinaryCompare) = 0 Then Exit Function
    NavyWriterMatches = (Len(term) = 0 Or InStr(1, NavyWriterValue(row, 0) & " " & NavyWriterValue(row, 1) & " " & NavyWriterValue(row, 3), term, vbTextCompare) > 0)
End Function
Public Function UXSelectNavyWriter() As Boolean
    Dim i As Long
    If IsNull(Forms!frmUXNavyWriter!lstCatalog) Then Exit Function
    i = CLng(Forms!frmUXNavyWriter!lstCatalog)
    If Not NavyWriterMatches(i, Nz(Forms!frmUXNavyWriter!cboRate, "ALL"), Nz(Forms!frmUXNavyWriter!txtSearch, "")) Then Exit Function
    Forms!frmUXNavyWriter!txtBody = NavyWriterValue(i, 1)
    Forms!frmUXNavyWriter!txtSource = NavyWriterValue(i, 2)
    Forms!frmUXNavyWriter!txtBody.Locked = (NavyWriterValue(i, 4) <> "Excerpt")
    Forms!frmUXNavyWriter!cmdImport.Enabled = (NavyWriterValue(i, 4) = "Excerpt")
    Forms!frmUXNavyWriter!cmdSource.Enabled = True
    Forms!frmUXNavyWriter!lblStatus.Caption = NavyWriterValue(i, 4) & " / Rate: " & Replace(NavyWriterValue(i, 3), "GENERAL", "General") & ". Adapt only to verified accomplishments; this is not policy."
    UXSelectNavyWriter = True
End Function
Public Function UXImportNavyWriter() As Boolean
    Dim i As Long, r As DAO.Recordset, db As DAO.Database, title As String, body As String, attribution As String, id As Long
    On Error GoTo Failed
    If IsNull(Forms!frmUXNavyWriter!lstCatalog) Then Err.Raise vbObjectError + 6920, "modUXNavyWriter.UXImportNavyWriter", "Select an example and review its wording first."
    i = CLng(Forms!frmUXNavyWriter!lstCatalog): title = "NavyWriter - " & NavyWriterValue(i, 0)
    If Not NavyWriterMatches(i, Nz(Forms!frmUXNavyWriter!cboRate, "ALL"), Nz(Forms!frmUXNavyWriter!txtSearch, "")) Then Err.Raise vbObjectError + 6920, "modUXNavyWriter.UXImportNavyWriter", "Selection is outside the current filter. Choose it again."
    If NavyWriterValue(i, 4) <> "Excerpt" Then Err.Raise vbObjectError + 6920, "modUXNavyWriter.UXImportNavyWriter", "This is a source collection. Open its source to review the full examples."
    body = Nz(Forms!frmUXNavyWriter!txtBody, ""): attribution = "NavyWriter excerpt [Rate: " & NavyWriterValue(i, 3) & "]: " & NavyWriterValue(i, 2)
    If Len(Trim$(body)) = 0 Then Err.Raise vbObjectError + 6920, "modUXNavyWriter.UXImportNavyWriter", "Enter reviewed narrative wording before adding it."
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT SnippetID,SnippetBody FROM NarrativeSnippet WHERE SnippetTitle=" & QuoteSQL(title) & " AND InStr(1,Attribution," & QuoteSQL(NavyWriterValue(i, 2)) & ",1)>0", dbOpenSnapshot)
    Do While Not r.EOF
        If StrComp(Nz(r!SnippetBody, ""), body, vbBinaryCompare) = 0 Then
            r.Close: Forms!frmUXNavyWriter!lblStatus.Caption = "This exact example is already in your narrative snippets.": UXImportNavyWriter = True: Exit Function
        End If
        r.MoveNext
    Loop
    r.Close
    id = SaveSnippet(0, title, body, attribution, "ALL")
    Forms!frmUXNavyWriter!lblStatus.Caption = "Added to your narrative snippets. Source attribution retained."
    UXImportNavyWriter = True: Exit Function
Failed:
    Forms!frmUXNavyWriter!lblStatus.Caption = Err.Description
End Function
Public Function UXOpenNavyWriterSource() As Boolean
    If IsNull(Forms!frmUXNavyWriter!lstCatalog) Then Exit Function
    Application.FollowHyperlink NavyWriterValue(CLng(Forms!frmUXNavyWriter!lstCatalog), 2)
    UXOpenNavyWriterSource = True
End Function
