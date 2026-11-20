Attribute VB_Name = "modUXGroupLifecycle"
Option Explicit
Public Sub ChangeSummaryGroup(ByVal id As Long, ByVal stamp As String, ByVal action As String, Optional ByVal title As String = "")
    Dim r As DAO.Recordset, db As DAO.Database, w As DAO.Workspace, op As Long, failure As Long, message As String, started As Boolean, oldTitle As String
    MajorAuthor
    If action <> "Rename" And action <> "Archive" And action <> "Recover" Then Err.Raise vbObjectError + 4900, "modUXGroupLifecycle.ChangeSummaryGroup", "Unknown group action."
    title = Trim$(title)
    If action = "Rename" And (Len(title) = 0 Or Len(title) > 120) Then Err.Raise vbObjectError + 4901, "modUXGroupLifecycle.ChangeSummaryGroup", "Enter a group name of 1 to 120 characters."
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0)
    On Error GoTo Failed
    w.BeginTrans: started = True
    If GroupStamp(id) <> stamp Then Err.Raise vbObjectError + 4902, "modUXGroupLifecycle.ChangeSummaryGroup", "This group changed. Refresh the list and try again."
    Set r = db.OpenRecordset("SELECT * FROM SummaryGroup WHERE GroupID=" & id, dbOpenDynaset)
    oldTitle = r!Title
    If action = "Rename" And r!IsArchived Then Err.Raise vbObjectError + 4904, "modUXGroupLifecycle.ChangeSummaryGroup", "Recover this group before renaming it."
    If action = "Archive" And r!IsArchived Then Err.Raise vbObjectError + 4904, "modUXGroupLifecycle.ChangeSummaryGroup", "Group is already archived."
    If action = "Recover" And Not r!IsArchived Then Err.Raise vbObjectError + 4904, "modUXGroupLifecycle.ChangeSummaryGroup", "Group is already active."
    If action = "Rename" And StrComp(title, oldTitle, vbBinaryCompare) = 0 Then r.Close: w.CommitTrans: Exit Sub
    op = BeginOperation(action & " summary group", "Group " & id & "; previous name: " & oldTitle)
    r.Edit
    If action = "Rename" Then r!Title = title Else r!IsArchived = (action = "Archive")
    r!VersionCounter = r!VersionCounter + 1: r.Update: r.Close
    FinishOperation op, "Completed", IIf(action = "Rename", "New name: " & title, action & " completed") & "; reports, members, defaults and history retained."
    w.CommitTrans: Call UXEditorContextChanged(): started = False: Exit Sub
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modUXGroupLifecycle", "ChangeSummaryGroup", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    If started Then SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Sub
Public Function UXGroupList() As Boolean
    Dim f As Access.Form, id As Long, filter As String, where As String, having As String, endYear As Long, i As Long, listed As Boolean
    Set f = Forms!frmUXGroups: id = Nz(f!lstGroups, 0): filter = Nz(f!uxArchiveFilter, "Active")
    If filter = "Active" Then where = "WHERE g.IsArchived=False "
    If filter = "Archived" Then where = "WHERE g.IsArchived=True "
    'Ending year is inclusive: a group shows when any report in it ends that year,
    'or, with no reports, when its default ending date does.
    If UXHasControl(f, "uxYearFilter") Then
        GroupYearChoices f
        If IsNumeric(Nz(f!uxYearFilter, "")) Then endYear = CLng(f!uxYearFilter)
    End If
    If endYear > 0 Then having = "HAVING Sum(IIf(Year(r.PeriodEnd)=" & endYear & ",1,0))>0 OR (Count(m.ReportID)=0 AND GroupDefaultEndYear(g.GroupID)=" & endYear & ") "
    f!lstGroups.RowSource = "SELECT g.GroupID,g.Title AS [Group],g.FormKind AS [Report type],g.Paygrade,Count(m.ReportID) AS Reports,IIf(g.IsArchived,'Archived','Active') AS [Archive state],First(IIf(Len(Nz(l.WorkflowState,''))=0,'Initialized',l.WorkflowState)) AS [Group status],GroupCloseout(g.GroupID,Count(r.ReportID),Count(r.PeriodEnd),Min(r.PeriodEnd),Max(r.PeriodEnd)) AS [Closeout date] FROM ((SummaryGroup AS g LEFT JOIN GroupLifecycle AS l ON g.GroupID=l.GroupID) LEFT JOIN GroupMember AS m ON g.GroupID=m.GroupID) LEFT JOIN Report AS r ON m.ReportID=r.ReportID " & where & "GROUP BY g.GroupID,g.Title,g.FormKind,g.Paygrade,g.IsArchived " & having & "ORDER BY g.Title"
    If id > 0 Then
        If DCount("*", "SummaryGroup", "GroupID=" & id) = 0 Then
            id = 0
        ElseIf filter = "Active" And Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & id), False) Then
            id = 0
        ElseIf filter = "Archived" And Not Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & id), True) Then
            id = 0
        End If
    End If
    If id > 0 And endYear > 0 Then
        For i = 0 To f!lstGroups.ListCount - 1
            If CStr(Nz(f!lstGroups.Column(0, i), "")) = CStr(id) Then listed = True: Exit For
        Next
        If Not listed Then id = 0
    End If
    On Error Resume Next
    f!lstGroups = id
    On Error GoTo 0
    Call UXGroupSelection
    f!lstGroups.ColumnCount = 8: f!lstGroups.ColumnWidths = "0;3200;1400;1200;1100;1600;2200;2000"
    UXResizeWorkspace f.Name: UXGroupList = True
End Function
Public Function GroupDefaultEndYear(ByVal groupID As Long) As Long
    Dim value As Variant
    value = UXEffectiveDefault(groupID, "PeriodEnd")
    If IsDate(value) Then GroupDefaultEndYear = Year(CDate(value))
End Function
'Year choices come from the reports in groups and the defaults of empty groups,
'newest first. A chosen year stays listed even when no group matches it now.
Private Sub GroupYearChoices(ByVal f As Access.Form)
    Dim years As Object, r As DAO.Recordset, chosen As String, choices As String, y As Long, k As Variant, newest As Long
    Set years = CreateObject("Scripting.Dictionary"): chosen = Nz(f!uxYearFilter, "All years")
    Set r = CurrentDb.OpenRecordset("SELECT DISTINCT Year(r.PeriodEnd) AS EndYear FROM GroupMember AS m INNER JOIN Report AS r ON m.ReportID=r.ReportID WHERE r.PeriodEnd IS NOT NULL", dbOpenSnapshot)
    Do While Not r.EOF
        years(CLng(r!EndYear)) = True: r.MoveNext
    Loop
    r.Close
    Set r = CurrentDb.OpenRecordset("SELECT GroupID FROM SummaryGroup WHERE GroupID NOT IN (SELECT GroupID FROM GroupMember)", dbOpenSnapshot)
    Do While Not r.EOF
        y = GroupDefaultEndYear(r!GroupID)
        If y > 0 Then years(y) = True
        r.MoveNext
    Loop
    r.Close
    If IsNumeric(chosen) Then years(CLng(chosen)) = True Else chosen = "All years"
    Do While years.Count > 0
        newest = 0
        For Each k In years.Keys
            If k > newest Then newest = k
        Next
        choices = choices & ";" & newest: years.Remove newest
    Loop
    choices = """All years""" & choices
    If f!uxYearFilter.RowSource <> choices Then f!uxYearFilter.RowSource = choices
    f!uxYearFilter = chosen
End Sub
Public Function GroupCloseout(ByVal groupID As Long, ByVal members As Long, ByVal dated As Long, ByVal firstDate As Variant, ByVal lastDate As Variant) As Variant
 Dim value As Variant
 If members = 0 Then
  value = UXEffectiveDefault(groupID, "PeriodEnd")
  If IsNull(value) Or Len(Nz(value, "")) = 0 Then GroupCloseout = Null Else GroupCloseout = Format$(CDate(value), "yyyy-mm-dd")
 ElseIf dated <> members Then
  GroupCloseout = "Incomplete"
 ElseIf firstDate <> lastDate Then
  GroupCloseout = "Mixed"
 Else
  GroupCloseout = Format$(lastDate, "yyyy-mm-dd")
 End If
End Function
Public Function UXGroupSelection() As Boolean
    Dim f As Access.Form, id As Long, archived As Boolean
    Set f = Forms!frmUXGroups: id = Nz(f!lstGroups, 0)
    If UXHasControl(f, "uxExportSelected") Then UXEnable f!uxExportSelected, (id > 0)
    If UXHasControl(f, "uxPDFSelected") Then UXEnable f!uxPDFSelected, (id > 0)
    If UXHasControl(f, "uxGroupState") Then
        If id > 0 Then f!uxGroupState = GroupWorkflow(id): f!uxStateStamp = GroupStamp(id) Else f!uxGroupState = Null: f!uxStateStamp = ""
    End If
    archived = Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & id), False)
    On Error Resume Next
    If Screen.ActiveControl.Parent.Name = f.Name Then
        If Screen.ActiveControl.Name = "uxArchive" Or Screen.ActiveControl.Name = "uxRecover" Or Screen.ActiveControl.Name = "uxRename" Or Screen.ActiveControl.Name = "cmdOpen" Then f!uxArchiveFilter.SetFocus
    End If
    On Error GoTo 0
    UXEnable f!cmdOpen, (id > 0): UXEnable f!uxRename, (id > 0 And Not archived)
    UXEnable f!uxArchive, (id > 0 And Not archived): UXEnable f!uxRecover, (id > 0 And archived)
    f!uxLifecycleStatus.Caption = IIf(id = 0, "Choose a group.", IIf(archived, "Archived group. Recover it to work in it.", ""))
    f!uxLifecycleStatus.ControlTipText = "Group status is separate from member review and official submission."
    UXGroupSelection = True
End Function
Public Function UXGroupAction(ByVal action As String) As Boolean
    Dim f As Access.Form, id As Long, title As String
    On Error GoTo Failed
    Set f = Forms!frmUXGroups: id = Nz(f!lstGroups, 0)
    If id = 0 Then Err.Raise vbObjectError + 4903, "modUXGroupLifecycle.UXGroupAction", "Choose a group first."
    title = Nz(DLookup("Title", "SummaryGroup", "GroupID=" & id), "")
    If action = "Rename" Then
        DoCmd.OpenForm "frmUXRenameGroup", , , , , acDialog, CStr(id)
    Else
        ChangeSummaryGroup id, GroupStamp(id), action
        Call UXGroupList
        f!uxLifecycleStatus.Caption = title & IIf(action = "Archive", " archived. Select Archived to recover it.", " recovered. Select Active to continue.")
        UXRefreshGroupWorkspace id
    End If
    UXGroupAction = True: Exit Function
Failed:
    Forms!frmUXGroups!uxLifecycleStatus.Caption = SwiftErrorText()
End Function
Public Function UXLoadRenameGroup() As Boolean
    With Forms!frmUXRenameGroup
        !groupID = CLng(.OpenArgs): !baseline = GroupStamp(!groupID)
        !txtName = DLookup("Title", "SummaryGroup", "GroupID=" & !groupID)
    End With
    UXLoadRenameGroup = True
End Function
Public Function UXSaveGroupName() As Boolean
    Dim id As Long
    On Error GoTo Failed
    With Forms!frmUXRenameGroup
        id = !groupID: ChangeSummaryGroup id, !baseline, "Rename", Nz(!txtName, "")
    End With
    DoCmd.Close acForm, "frmUXRenameGroup", acSaveNo
    Call UXGroupList
    UXRefreshGroupWorkspace id
    UXSaveGroupName = True: Exit Function
Failed:
    Forms!frmUXRenameGroup!lblStatus.Caption = SwiftErrorText()
End Function
Public Sub UXRefreshGroupWorkspace(ByVal id As Long)
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then
        If Nz(Forms!frmUXComparison!uxGroupContext, 0) = id And Nz(Forms!frmUXComparison!txtMeeting, 0) = 0 Then
            UXShowGroupMembers id, True
            Exit Sub
        End If
    End If
    If CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then
        If Nz(Forms!frmMajorWorkbench!cboGroup, 0) = id Then UXFocusedGroupMode Nz(Forms!frmMajorWorkbench!uxMode, "Members")
    End If
End Sub
Public Sub UXGroupLocks(ByVal f As Access.Form, ByVal id As Long)
    Dim name As Variant, archived As Boolean
    archived = Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & id), True)
    For Each name In Array("uxMode1", "uxMode3", "uxMembers", "uxBulk", "uxRSCA", "uxComparison", "uxPolicyReview", "uxAddGroupReport", "uxDefaultsEditor", "cmdStartRoute", "cmdCreateRoute", "cmdAcceptStep", "cmdReturnStep", "cmdCancelRoute", "cmdApproveRoute", "cmdReopenApproval", "cmdApplyBatch", "cmdPreviewUndo")
        If UXHasControl(f, CStr(name)) Then UXEnable f.Controls(name), Not archived
    Next
    If archived Then f!lblStatus.Caption = "Archived group: read only. Recover it from Summary groups to manage it."
End Sub
