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
    r.Close: Call RefreshNarrativeLabel(): Call CheckpointEditor()
    Forms!frmEditor.SetFocus: InsertNarrativeSnippet = True
End Function
Public Function LoadSnippets() As Boolean
    Dim term As String
    term = Replace(Left$(Nz(Forms!frmSnippets!txtSearch, ""), 255), "'", "''")
    Forms!frmSnippets!lstSnippets.RowSource = "SELECT SnippetID,SnippetTitle,FormKind,Attribution FROM NarrativeSnippet WHERE InStr(1,SnippetTitle,'" & term & "',1)>0 OR InStr(1,SnippetBody,'" & term & "',1)>0 ORDER BY SnippetTitle"
    Forms!frmSnippets!lstSnippets.Requery: LoadSnippets = True
End Function
Public Function SelectSnippet() As Boolean
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT * FROM NarrativeSnippet WHERE SnippetID=" & Nz(Forms!frmSnippets!lstSnippets, 0), dbOpenSnapshot)
    If Not r.EOF Then
        Forms!frmSnippets!txtSnippetID = r!SnippetID: Forms!frmSnippets!txtTitle = r!SnippetTitle
        Forms!frmSnippets!txtBody = r!SnippetBody: Forms!frmSnippets!txtAttribution = r!Attribution: Forms!frmSnippets!cboKind = r!FormKind
    End If
    r.Close: SelectSnippet = True
End Function
Public Function NewSnippetUI() As Boolean
    Forms!frmSnippets!txtSnippetID = 0: Forms!frmSnippets!txtTitle = "": Forms!frmSnippets!txtBody = "": Forms!frmSnippets!txtAttribution = "": Forms!frmSnippets!cboKind = "ALL"
    NewSnippetUI = True
End Function
Public Function SaveSnippetUI() As Boolean
    On Error GoTo Failed
    With Forms!frmSnippets
        !txtSnippetID = SaveSnippet(Nz(!txtSnippetID, 0), Nz(!txtTitle, ""), Nz(!txtBody, ""), Nz(!txtAttribution, ""), Nz(!cboKind, "ALL"))
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
    DoCmd.Close acForm, "frmSnippets", acSaveNo
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Forms!frmEditor.SetFocus Else Call OpenScreen("frmHome")
    CloseSnippetsUI = True
End Function
