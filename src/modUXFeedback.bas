Attribute VB_Name = "modUXFeedback"
Option Explicit
Private snippetReport As Long
Private gridReport As Long
Private catalogFromSnippets As Boolean
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
    On Error GoTo Failed
    SwitchContributor CLng(Nz(Forms!frmHome!uxContributor, 0))
    LoadHome
    UXLoadContext
    UXOverviewContributor = True: Exit Function
Failed:
    Forms!frmHome!uxContributor = MajorAuthor()
    Forms!frmHome!uxContributorStatus.Caption = Err.Description
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
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!cmdClose.Caption = "Back to group members"
    UXGridOpenReport = True
End Function
Public Function UXFeedbackCloseEditor() As Boolean
    Dim id As Long
    id = Nz(Forms!frmEditor!txtReportID, 0)
    If Not UXLeaveEditor() Then Exit Function
    If id = gridReport And CurrentProject.AllForms("frmUXComparison").IsLoaded Then
        Forms!frmUXComparison.SetFocus
        UXGridCurrent
    End If
    gridReport = 0: UXFeedbackCloseEditor = True
End Function

Public Function UXWorkContributor() As Boolean
    On Error GoTo Failed
    SwitchContributor CLng(Nz(Forms!frmMyWork!uxContributor, 0))
    UXLoadContext
    UXLoadWorkQueue
    UXWorkContributor = True: Exit Function
Failed:
    Forms!frmMyWork!uxContributor = MajorAuthor()
    Forms!frmMyWork!uxWorkStatus.Caption = Err.Description
End Function
Public Function UXGroupRowOpen() As Boolean
    Dim id As Long
    id = Nz(Forms!frmMajorWorkbench!lstReports, 0)
    If id = 0 Then Exit Function
    If Not UXLeaveEditor() Then Exit Function
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
