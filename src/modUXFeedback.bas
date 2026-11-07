Attribute VB_Name = "modUXFeedback"
Option Explicit
Private snippetReport As Long
Private gridReport As Long
Private catalogFromSnippets As Boolean
Private siblingGroup As Long, siblingOrder As Collection
Public Function UXOpenReportSnippets() As Boolean
    On Error GoTo Failed
    UXFlushSurface
    snippetReport = Nz(Forms!frmEditor!txtReportID, 0)
    DoCmd.OpenForm "frmSnippets"
    UXOpenReportSnippets = True: Exit Function
Failed:
    UXShowSaveError Err.Description
End Function
Public Function UXReturnSnippets() As Boolean
    DoCmd.Close acForm, "frmSnippets", acSaveNo
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        If Nz(Forms!frmEditor!txtReportID, 0) = snippetReport Then
            Forms!frmEditor.SetFocus
            snippetReport = 0: UXReturnSnippets = True: Exit Function
        End If
    End If
    snippetReport = 0
    DoCmd.OpenForm "frmUXResources"
    UXReturnSnippets = True
End Function
Public Function UXResetSnippetOrigin() As Boolean
    snippetReport = -1: UXResetSnippetOrigin = True
End Function
Public Function UXOverviewContributor() As Boolean
    Dim message As String
    On Error GoTo Failed
    SwitchContributor CLng(Nz(Forms!frmHome!uxContributor, 0))
    LoadHome
    UXLoadContext
    UXOverviewContributor = True: Exit Function
Failed:
    'The contributor check sets its own handler, which clears Err; with no
    'session contributor, MajorAuthor would raise inside this handler.
    message = Err.Description
    If HasActiveContributor() Then Forms!frmHome!uxContributor = MajorAuthor() Else Forms!frmHome!uxContributor = Null
    Forms!frmHome!uxContributorStatus.Caption = message
End Function
Public Function UXUndoPending() As Boolean
    Dim id As Long
    On Error GoTo Failed
    id = Nz(Forms!frmEditor!txtReportID, 0)
    If Not UXLeaveEditor("Discard") Then Exit Function
    If id > 0 Then DoCmd.OpenForm "frmEditor", , , , , , CStr(id) Else DoCmd.OpenForm "frmLibrary"
    UXUndoPending = True: Exit Function
Failed:
    If CurrentProject.AllForms("frmEditor").IsLoaded Then UXShowSaveError "Unable to reload saved report: " & Err.Description
End Function
Public Function UXHideAccessObjects() As Boolean
    On Error GoTo Failed
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    DoCmd.SelectObject acTable, , True
    DoCmd.RunCommand acCmdWindowHide
    UXHideAccessObjects = True
Failed:
End Function
Public Function UXTableITA(ByVal id As Long) As String
    Dim average As Variant
    If Nz(DLookup("ObservationStatus", "Report", "ReportID=" & id), "") = "NOB" Then UXTableITA = "NOB": Exit Function
    average = NumericTraitAverage(CurrentRevision(id))
    If IsNull(average) Then UXTableITA = "Unavailable" Else UXTableITA = Format$(average, "0.00")
End Function
Public Function UXTableGroups(ByVal id As Long) As String
    Dim r As DAO.Recordset, db As DAO.Database
    Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT g.Title FROM SummaryGroup AS g INNER JOIN GroupMember AS m ON g.GroupID=m.GroupID WHERE m.ReportID=" & id & " ORDER BY g.Title,g.GroupID", dbOpenSnapshot)
    Do While Not r.EOF
        If Len(UXTableGroups) > 0 Then UXTableGroups = UXTableGroups & "; "
        UXTableGroups = UXTableGroups & r!Title: r.MoveNext
    Loop
    r.Close
    If Len(UXTableGroups) = 0 Then UXTableGroups = "Ungrouped"
End Function
Public Function UXListGroupDefaults() As Boolean
    Dim id As Long
    id = Nz(Forms!frmUXGroups!lstGroups, 0)
    If id = 0 Then Exit Function
    UXListGroupDefaults = UXOpenGroupDefaults(id)
End Function
Public Function UXOpenGroupMembers() As Boolean
    UXOpenGroupMembers = UXOpenGroup()
End Function
Public Function UXGridOpenReport() As Boolean
    Dim id As Long
    id = Nz(Forms!frmUXComparison!uxTraitGrid.Form!ReportID, 0)
    If id = 0 Then Exit Function
    If Not UXLeaveEditor() Then Exit Function
    gridReport = id
    UXCaptureGridOrder
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!cmdClose.Caption = "Back to group members"
    Forms!frmEditor!uxReturnGroup = Nz(Forms!frmUXComparison!uxGroupContext, 0)
    Call UXEditorSiblingState: UXDirectorySourceInvalidate: UXDirectorySourceStyle
    UXGridOpenReport = True
End Function
Public Function UXFeedbackCloseEditor() As Boolean
    Dim id As Long, returnGroup As Long
    id = Nz(Forms!frmEditor!txtReportID, 0)
    returnGroup = Nz(Forms!frmEditor!uxReturnGroup, 0)
    If Not UXLeaveEditor() Then Exit Function
    If returnGroup > 0 Then
        UXShowGroupMembers returnGroup, True
    ElseIf id = gridReport And CurrentProject.AllForms("frmUXComparison").IsLoaded Then
        Forms!frmUXComparison.SetFocus
        UXGridCurrent True
    End If
    gridReport = 0: UXFeedbackCloseEditor = True
End Function

Public Function UXWorkContributor() As Boolean
    Dim message As String
    On Error GoTo Failed
    SwitchContributor CLng(Nz(Forms!frmMyWork!uxContributor, 0))
    UXLoadContext
    UXLoadWorkQueue
    UXWorkContributor = True: Exit Function
Failed:
    'Capture before the contributor check, whose own handler clears Err.
    message = Err.Description
    If HasActiveContributor() Then Forms!frmMyWork!uxContributor = MajorAuthor() Else Forms!frmMyWork!uxContributor = Null
    Forms!frmMyWork!uxWorkStatus.Caption = message
End Function
Public Function UXGroupRowOpen() As Boolean
    Dim id As Long
    id = Nz(Forms!frmMajorWorkbench!lstReports, 0)
    If id = 0 Then Exit Function
    If Not UXLeaveEditor() Then Exit Function
    UXCaptureListOrder Nz(Forms!frmMajorWorkbench!cboGroup, 0), Forms!frmMajorWorkbench!lstReports
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    UXGroupRowOpen = True
End Function

Public Function UXOpenSnippetSources() As Boolean
    catalogFromSnippets = True
    DoCmd.OpenForm "frmUXNavyWriter"
    Forms!frmUXNavyWriter!cmdBack.Caption = "Back to narrative snippets"
    UXOpenSnippetSources = True
End Function
Public Function UXReturnNarrativeSources() As Boolean
    DoCmd.Close acForm, "frmUXNavyWriter", acSaveNo
    If catalogFromSnippets And CurrentProject.AllForms("frmSnippets").IsLoaded Then
        Forms!frmSnippets.SetFocus: LoadSnippets
    Else
        DoCmd.OpenForm "frmUXResources"
    End If
    catalogFromSnippets = False: UXReturnNarrativeSources = True
End Function

'Previous/Next follow the group view as the contributor last saw it: the order of
'the Members table (its heading sort) or list when a report was opened from it.
Public Sub UXCaptureGridOrder()
    Dim rows As DAO.Recordset, ids As New Collection
    On Error GoTo Finished
    Set rows = Forms!frmUXComparison!uxTraitGrid.Form.RecordsetClone
    If Not (rows.BOF And rows.EOF) Then rows.MoveFirst
    Do While Not rows.EOF
        If Nz(rows!ReportID, 0) > 0 Then ids.Add CLng(rows!ReportID)
        rows.MoveNext
    Loop
    siblingGroup = Nz(Forms!frmUXComparison!uxGroupContext, 0): Set siblingOrder = ids
Finished:
End Sub
Public Sub UXCaptureListOrder(ByVal groupID As Long, ByVal list As Access.Control)
    Dim ids As New Collection, i As Long
    For i = IIf(list.ColumnHeads, 1, 0) To list.ListCount - 1
        If Val(Nz(list.ItemData(i), 0)) > 0 Then ids.Add CLng(list.ItemData(i))
    Next
    siblingGroup = groupID: Set siblingOrder = ids
End Sub
Private Function EditorGroupReports(ByVal groupID As Long) As Collection
    Dim ids As New Collection, r As DAO.Recordset, alphabetical As New Collection, active As Object, placed As Object, item As Variant
    If groupID > 0 And groupID = siblingGroup And Not siblingOrder Is Nothing Then
        'Captured order first, then any active member added since, alphabetically.
        Set active = CreateObject("Scripting.Dictionary"): Set placed = CreateObject("Scripting.Dictionary")
        Set r = CurrentDb.OpenRecordset("SELECT r.ReportID FROM Report AS r INNER JOIN GroupMember AS g ON r.ReportID=g.ReportID WHERE g.GroupID=" & groupID & " AND r.IsArchived=False ORDER BY UXIdentity(r.OfficialName,r.MemberName),r.ReportID", dbOpenSnapshot)
        Do While Not r.EOF: alphabetical.Add CLng(r!ReportID): active(CStr(r!ReportID)) = True: r.MoveNext: Loop
        r.Close
        For Each item In siblingOrder
            If active.Exists(CStr(item)) And Not placed.Exists(CStr(item)) Then ids.Add CLng(item): placed(CStr(item)) = True
        Next
        For Each item In alphabetical
            If Not placed.Exists(CStr(item)) Then ids.Add CLng(item)
        Next
        Set EditorGroupReports = ids: Exit Function
    End If
    If groupID > 0 Then
        Set r = CurrentDb.OpenRecordset("SELECT r.ReportID FROM Report AS r INNER JOIN GroupMember AS g ON r.ReportID=g.ReportID WHERE g.GroupID=" & groupID & " AND r.IsArchived=False ORDER BY UXIdentity(r.OfficialName,r.MemberName),r.ReportID", dbOpenSnapshot)
        Do While Not r.EOF: ids.Add CLng(r!ReportID): r.MoveNext: Loop
        r.Close
    End If
    Set EditorGroupReports = ids
End Function
Public Function UXEditorSiblingState() As Boolean
    Dim f As Access.Form, ids As Collection, groupID As Long, i As Long, position As Long
    Set f = Forms!frmEditor
    groupID = Nz(f!uxReturnGroup, 0): If groupID = 0 Then groupID = Nz(f!uxGroup, 0)
    Set ids = EditorGroupReports(groupID)
    For i = 1 To ids.Count: If ids(i) = Nz(f!txtReportID, 0) Then position = i
    Next
    f!uxPreviousReport.Visible = (position > 0): f!uxNextReport.Visible = (position > 0)
    f!uxPreviousReport.Enabled = (position > 1): f!uxNextReport.Enabled = (position > 0 And position < ids.Count)
    f!uxPreviousReport.ControlTipText = "Previous report in this summary group, in the order the group's Members view showed."
    f!uxNextReport.ControlTipText = "Next report in this summary group, in the order the group's Members view showed."
    UXEditorSiblingState = True
End Function
Public Function UXEditorSiblingMove(ByVal direction As Long, Optional ByVal decision As String = "") As Boolean
    Dim f As Access.Form, ids As Collection, groupID As Long, current As Long, target As Long, i As Long, position As Long
    On Error GoTo Failed
    If direction <> -1 And direction <> 1 Then Err.Raise 5, , "Choose Previous or Next."
    Set f = Forms!frmEditor: current = Nz(f!txtReportID, 0)
    groupID = Nz(f!uxReturnGroup, 0): If groupID = 0 Then groupID = Nz(f!uxGroup, 0)
    Set ids = EditorGroupReports(groupID)
    For i = 1 To ids.Count: If ids(i) = current Then position = i
    Next
    If position = 0 Or position + direction < 1 Or position + direction > ids.Count Then Exit Function
    target = ids(position + direction)
    If Not UXLeaveEditor(decision) Then Exit Function
    If DCount("*", "GroupMember", "GroupID=" & groupID & " AND ReportID=" & target) = 0 Then Err.Raise 5, , "Group membership changed. Open the group again."
    DoCmd.OpenForm "frmEditor", , , , , , CStr(target)
    Forms!frmEditor!uxReturnGroup = groupID
    Forms!frmEditor!cmdClose.Caption = "Back to group members"
    Call UXEditorContextChanged: UXDirectorySourceInvalidate
    UXRefreshEditor
    UXEditorSiblingMove = True: Exit Function
Failed:
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Forms!frmEditor!uxError.Caption = "Report navigation: " & Err.Description
End Function
