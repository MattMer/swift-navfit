Attribute VB_Name = "modSnippets"
Option Compare Database
Option Explicit
Public Function SaveSnippet(ByVal snippetID As Long, ByVal title As String, ByVal body As String, ByVal attribution As String, ByVal kind As String) As Long
    Dim r As DAO.Recordset
    If Len(Trim$(title)) = 0 Or Len(title) > 120 Or Len(Trim$(body)) = 0 Or Len(Trim$(attribution)) = 0 Or Len(attribution) > 255 Then Err.Raise vbObjectError + 1200, , "Provide a title (120 max), text and attribution (255 max)."
    If kind <> "ALL" And kind <> "EVAL" And kind <> "CHIEFEVAL" And kind <> "FITREP" Then Err.Raise vbObjectError + 1201, , "Choose a supported report type or ALL."
    Set r = CurrentDb.OpenRecordset("SELECT * FROM NarrativeSnippet WHERE SnippetID=" & snippetID, dbOpenDynaset)
    If snippetID = 0 Then
        r.AddNew: r!SnippetGuid = NewIdentity(): r!CreatedBy = DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1")
    Else
        If r.EOF Then Err.Raise vbObjectError + 1202, , "Snippet missing."
        r.Edit
    End If
    r!SnippetTitle = title: r!SnippetBody = body: r!Attribution = attribution: r!FormKind = kind
    r!ModifiedAt = Now(): r.Update: r.Bookmark = r.LastModified: SaveSnippet = r!SnippetID: r.Close
End Function
Public Function InsertNarrativeSnippet(ByVal snippetID As Long) As Boolean
    Dim r As DAO.Recordset, text As String, kind As String
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 1203, , "Open the target draft editor first."
    If Not Forms!frmEditor!cmdSave.Enabled Then Err.Raise vbObjectError + 1204, , "Recover the archived report before editing."
    Set r = CurrentDb.OpenRecordset("SELECT * FROM NarrativeSnippet WHERE SnippetID=" & snippetID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 1205, , "Select a snippet."
    kind = Nz(Forms!frmEditor!cboKind, "")
    If r!FormKind <> "ALL" And r!FormKind <> kind Then Err.Raise vbObjectError + 1206, , "Snippet is intended for a different report type."
    text = Nz(Forms!frmEditor!txtOfficialComments, "")
    If Len(text) > 0 Then text = text & vbCrLf
    Forms!frmEditor!txtOfficialComments = text & r!SnippetBody
    r.Close: Call RefreshNarrativeLabel(): Call CheckpointEditor(): Call UXRefreshEditor()
    Forms!frmEditor.SetFocus: InsertNarrativeSnippet = True
End Function
Public Function SnippetRate(ByVal attribution As Variant) As String
    Dim text As String, start As Long, finish As Long
    text = Nz(attribution, ""): start = InStr(1, text, "[Rate: ", vbBinaryCompare)
    If start = 0 Then SnippetRate = "UNSPECIFIED": Exit Function
    finish = InStr(start, text, "]")
    If finish = 0 Then SnippetRate = "UNSPECIFIED": Exit Function
    SnippetRate = Mid$(text, start + 7, finish - start - 7)
End Function
Public Function SnippetRateAttribution(ByVal attribution As String, ByVal rate As String) As String
    Dim start As Long, finish As Long
    start = InStr(1, attribution, "[Rate: ", vbBinaryCompare)
    If start > 0 Then
        finish = InStr(start, attribution, "]")
        If finish > 0 Then attribution = Trim$(Left$(attribution, start - 1) & Mid$(attribution, finish + 1))
    End If
    If rate <> "UNSPECIFIED" Then attribution = attribution & " [Rate: " & rate & "]"
    SnippetRateAttribution = attribution
End Function
Public Function LoadSnippets() As Boolean
    Dim term As String, rate As Variant, where As String, f As Access.Form
    Set f = Forms!frmSnippets
    If Len(f!cboRateFilter.RowSource) = 0 Then
        f!cboRateFilter.RowSourceType = "Value List": f!cboRateFilter.RowSource = "ALL;All rates;UNSPECIFIED;Unassigned / legacy;GENERAL;General / any rate"
        f!cboSnippetRate.RowSourceType = "Value List": f!cboSnippetRate.RowSource = "UNSPECIFIED;Unassigned / legacy;GENERAL;General / any rate"
        For Each rate In NavyWriterRates()
            f!cboRateFilter.AddItem CStr(rate) & ";" & CStr(rate): f!cboSnippetRate.AddItem CStr(rate) & ";" & CStr(rate)
        Next
        f!cboRateFilter = "ALL": f!cboSnippetRate = "UNSPECIFIED"
    End If
    term = Replace(Left$(Nz(f!txtSearch, ""), 255), "'", "''")
    where = "(InStr(1,SnippetTitle,'" & term & "',1)>0 OR InStr(1,SnippetBody,'" & term & "',1)>0)"
    If Nz(f!cboRateFilter, "ALL") <> "ALL" Then where = where & " AND SnippetRate([Attribution])=" & QuoteSQL(CStr(f!cboRateFilter))
    f!lstSnippets.RowSource = "SELECT SnippetID,SnippetTitle,FormKind,Attribution FROM NarrativeSnippet WHERE " & where & " ORDER BY SnippetTitle"
    f!lstSnippets.Requery: LoadSnippets = True
End Function
Public Function SelectSnippet() As Boolean
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT * FROM NarrativeSnippet WHERE SnippetID=" & Nz(Forms!frmSnippets!lstSnippets, 0), dbOpenSnapshot)
    If Not r.EOF Then
        Forms!frmSnippets!txtSnippetID = r!SnippetID: Forms!frmSnippets!txtTitle = r!SnippetTitle
        Forms!frmSnippets!txtBody = r!SnippetBody: Forms!frmSnippets!txtAttribution = r!Attribution: Forms!frmSnippets!cboKind = r!FormKind: Forms!frmSnippets!cboSnippetRate = SnippetRate(r!Attribution)
    End If
    r.Close: SelectSnippet = True
End Function
Public Function NewSnippetUI() As Boolean
    Forms!frmSnippets!txtSnippetID = 0: Forms!frmSnippets!txtTitle = "": Forms!frmSnippets!txtBody = "": Forms!frmSnippets!txtAttribution = "": Forms!frmSnippets!cboKind = "ALL"
    Forms!frmSnippets!cboSnippetRate = "UNSPECIFIED"
    NewSnippetUI = True
End Function
Public Function SaveSnippetUI() As Boolean
    On Error GoTo Failed
    With Forms!frmSnippets
        !txtSnippetID = SaveSnippet(Nz(!txtSnippetID, 0), Nz(!txtTitle, ""), Nz(!txtBody, ""), SnippetRateAttribution(Nz(!txtAttribution, ""), Nz(!cboSnippetRate, "UNSPECIFIED")), Nz(!cboKind, "ALL"))
    End With
    Call LoadSnippets(): SaveSnippetUI = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Save snippet"
End Function
Public Function InsertSnippetUI() As Boolean
    On Error GoTo Failed
    If MsgBox("Append this snippet to the open draft's official narrative? Review it for this member before saving. Existing text is retained.", vbOKCancel + vbQuestion) <> vbOK Then Exit Function
    InsertSnippetUI = InsertNarrativeSnippet(Nz(Forms!frmSnippets!txtSnippetID, 0)): Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Insert snippet"
End Function
Public Function CloseSnippetsUI() As Boolean
    CloseSnippetsUI = UXReturnSnippets()
End Function
