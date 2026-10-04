Attribute VB_Name = "modUXDashboard"
Option Compare Binary
Option Explicit
Public Function UXMemberLabel(ByVal reportID As Long) As String
    UXMemberLabel = Nz(DLookup("MemberName", "Report", "ReportID=" & reportID), "")
    If Len(Trim$(UXMemberLabel)) = 0 Then UXMemberLabel = Nz(DLookup("OfficialName", "Report", "ReportID=" & reportID), "")
    If Len(Trim$(UXMemberLabel)) = 0 Then UXMemberLabel = "Unnamed member"
End Function
Public Function UXReportState(ByVal reportID As Long) As String
    Dim revision As Long
    If Nz(DLookup("IsArchived", "Report", "ReportID=" & reportID), False) Then UXReportState = "Archived": Exit Function
    If Nz(DLookup("WorkflowState", "Report", "ReportID=" & reportID), "Draft") = "Approved" Then UXReportState = "Local approval": Exit Function
    revision = Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & reportID), 0)
    If DCount("*", "ReviewTask", "ReportID=" & reportID & " AND TaskState='Pending' AND RevisionID<>" & revision) > 0 Then UXReportState = "Review needs restart": Exit Function
    If DCount("*", "RouteRun", "ReportID=" & reportID & " AND RunState In ('Pending','Accepted') AND RevisionID<>" & revision) > 0 Then UXReportState = "Review needs restart": Exit Function
    If DCount("*", "ReviewTask", "ReportID=" & reportID & " AND TaskState='Pending'") > 0 Or DCount("*", "RouteRun", "ReportID=" & reportID & " AND RunState='Pending'") > 0 Then UXReportState = "In review": Exit Function
    If DCount("*", "ReviewTask", "ReportID=" & reportID & " AND TaskState='Returned' AND RevisionID=" & revision) > 0 Then UXReportState = "Returned": Exit Function
    UXReportState = "Draft"
End Function
Public Function UXGroupAverage(ByVal groupID As Long) As Variant
    Dim r As DAO.Recordset, key As Variant, value As Variant, total As Double, marks As Long
    Set r = CurrentDb.OpenRecordset("SELECT r.CurrentRevisionID,r.ObservationStatus FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & groupID & " AND r.IsArchived=False", dbOpenSnapshot)
    Do While Not r.EOF
        If Nz(r!ObservationStatus, "") <> "NOB" Then
            For Each key In TraitKeys()
                value = MajorRevisionValue(Nz(r!CurrentRevisionID, 0), CStr(key))
                Select Case Nz(value, "")
                    Case "1.0", "2.0", "3.0", "4.0", "5.0": total = total + CDbl(value): marks = marks + 1
                End Select
            Next
        End If
        r.MoveNext
    Loop
    r.Close
    If marks > 0 Then UXGroupAverage = total / marks Else UXGroupAverage = Null
End Function
Public Function UXNewGroup() As Boolean
    If Not UXLeaveEditor() Then Exit Function
    DoCmd.OpenForm "frmUXNewGroup", , , , , acDialog
    UXNewGroup = True
End Function
Public Function UXCreateGroupUI() As Boolean
    Dim id As Long
    On Error GoTo Failed
    With Forms!frmUXNewGroup
        id = CreateSummaryGroup(Nz(!txtTitle, ""), Nz(!cboKind, ""), Nz(!cboPaygrade, ""))
    End With
    DoCmd.Close acForm, "frmUXNewGroup", acSaveNo
    DoCmd.OpenForm "frmMajorWorkbench"
    Forms!frmMajorWorkbench!cboGroup = id
    UXWorkbenchMode "Members"
    If CurrentProject.AllForms("frmUXGroups").IsLoaded Then UXLoadGroups
    UXCreateGroupUI = True: Exit Function
Failed:
    Forms!frmUXNewGroup!lblStatus.Caption = Err.Description
End Function
Public Function UXOpenGroupReport() As Boolean
    Dim id As Long, selected As Variant
    For Each selected In Forms!frmMajorWorkbench!lstReports.ItemsSelected
        id = Forms!frmMajorWorkbench!lstReports.ItemData(selected): Exit For
    Next
    If id = 0 Then Exit Function
    If Not UXLeaveEditor() Then Exit Function
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    UXOpenGroupReport = True
End Function
Private Sub GroupPosition(ByVal name As String, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long)
    With Forms!frmMajorWorkbench.Controls(name)
        .Left = x: .Top = y: .Width = width: .Height = height: .Visible = True
    End With
End Sub
Public Function UXFocusedGroupMode(ByVal mode As String) As Boolean
    Dim f As Access.Form, c As Access.Control, id As Long, grade As String, average As Variant, psg As Object, name As Variant, i As Long, modes As Variant
    Set f = Forms!frmMajorWorkbench
    If Not UXHasControl(f, "uxGroupTitle") Then Exit Function
    id = Nz(f!cboGroup, 0)
    If id = 0 Then f.ScrollBars = 2: f.Section(acDetail).Height = 12500: Exit Function
    f.ScrollBars = 0: f.Section(acDetail).Height = 8600
    f!uxGroupTitle.Caption = Nz(DLookup("Title", "SummaryGroup", "GroupID=" & id), "Summary group")
    average = UXGroupAverage(id): Set psg = RSGroupPSG(id)
    f!uxGroupMeta.Caption = Nz(DLookup("FormKind", "SummaryGroup", "GroupID=" & id), "") & " | " & Nz(DLookup("Paygrade", "SummaryGroup", "GroupID=" & id), "") & " | " & DCount("*", "GroupMember", "GroupID=" & id) & " members | Trait average: " & IIf(IsNull(average), "Unavailable", Format$(Nz(average, 0), "0.00")) & " | PSG RSCA: " & IIf(IsNull(psg("Value")), "Unavailable", Format$(Nz(psg("Value"), 0), "0.00") & IIf(psg("Estimated"), " (estimate)", ""))
    For Each c In f.Controls
        If Left$(c.Tag, 4) = "WBX:" Then c.Visible = False
    Next
    f!lblStatus.Visible = True: f!uxBack.Visible = True
    If mode = "Members" Or mode = "Review" Then
        GroupPosition "lstReports", 2850, 3000, 10450, 2900
        f!lstReports.RowSource = "SELECT r.ReportID,r.MemberName AS Member,r.FormKind AS [Report type],r.PeriodEnd AS [Period to],UXReportState(r.ReportID) AS [Local status] FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & id & " ORDER BY r.MemberName"
        f!lstReports.ColumnCount = 5: f!lstReports.ColumnWidths = "0;3800;1500;2200;2800": f!lstReports.Requery
        f!uxOpenGroupReport.Visible = True
        If mode = "Members" Then f!uxMembers.Visible = True: f!uxBulk.Visible = True: f!uxRSCA.Visible = True
    ElseIf mode = "Comparison" Then
        f!uxComparison.Visible = True
        f!lblStatus.Caption = "Compare saved trait marks and stage proposed edits. A comparison baseline is captured automatically."
    ElseIf mode = "Defaults" Then
        f!cboScope = "Group"
        GroupPosition "cboField", 2850, 3000, 4800, 450
        GroupPosition "txtValue", 8100, 3000, 5200, 450
        GroupPosition "chkNull", 2850, 3650, 320, 320
        GroupPosition "cmdSaveDefault", 2850, 4200, 4800, 500
        GroupPosition "cmdSuggestDefault", 8100, 4200, 5200, 500
        GroupPosition "txtDetail", 2850, 4900, 10450, 1700
        f!uxBulk.Visible = True
        f!lblStatus.Caption = "Defaults are suggestions. Edit shared report blocks to preview and apply changes to members."
    ElseIf mode = "History" Then
        GroupPosition "cboBatch", 2850, 3000, 10450, 450
        GroupPosition "cmdShowBatch", 2850, 3600, 3200, 450
        GroupPosition "cmdPreviewUndo", 6400, 3600, 3200, 450
        GroupPosition "cmdApplyBatch", 9900, 3600, 3400, 450
        GroupPosition "lstDetails", 2850, 4200, 10450, 1200
        GroupPosition "txtDetail", 2850, 5550, 10450, 1250
    End If
    modes = Array("Members", "Defaults", "Review", "Comparison", "History")
    For i = 0 To 4
        f.Controls("uxMode" & i).FontBold = (modes(i) = mode)
    Next
    UXFocusedGroupMode = True
End Function
Public Function UXComparisonPage(ByVal page As String) As Boolean
    Dim f As Access.Form, c As Access.Control, preview As Boolean
    Set f = Forms!frmUXComparison: preview = (page = "Preview")
    For Each c In f.Controls
        If c.Name Like "cboTrait##" Or c.Name Like "lblTrait##" Or c.Name = "lstTraits" Or c.Name = "uxMember" Or c.Name = "uxITA" Or c.Name = "cboPromotion" Or c.Name = "cmdStage" Then c.Visible = Not preview
    Next
    f!lstProposals.Visible = preview: f!cmdApply.Visible = preview
    f!uxMarksPage.FontBold = Not preview: f!uxPreviewPage.FontBold = preview
    UXComparisonPage = True
End Function
Public Function UXMyWorkMode(ByVal mode As String) As Boolean
    Dim f As Access.Form, c As Access.Control, names As Variant, name As Variant
    Set f = Forms!frmMyWork
    If Not UXHasControl(f, "uxWorkTab0") Then Exit Function
    f!uxWorkTab0.SetFocus
    For Each c In f.Controls
        If Left$(c.Tag, 4) = "MWX:" Then c.Visible = False
    Next
    Select Case mode
        Case "Contributors": names = Array("cboContributor", "cmdSwitch", "txtContributor", "cmdAdd", "lblScope")
        Case "Assign": names = Array("uxAssignReport", "uxAssignCaption", "lblAssignment", "lblReviewer", "cboReviewer", "lblDue", "txtDue", "lblNotes", "txtNotes", "cmdAssign")
        Case Else: names = Array("chkAll", "lblAll", "cmdRefresh", "lstTasks", "txtHistory", "lblNotes", "txtNotes", "cmdAccept", "cmdReturn", "cmdCancel", "cmdDraft")
    End Select
    For Each name In names: f.Controls(CStr(name)).Visible = True: Next
    f!lblNotes.Caption = IIf(mode = "Assign", "Assignment instructions", "Decision note (required when returning a report)")
    If mode = "Assign" Then
        f!uxAssignReport.RowSource = "SELECT ReportID,MemberName & ' | ' & FormKind & ' | ' & Nz(Format(PeriodEnd,'yyyy-mm-dd'),'No period') AS Description FROM Report WHERE IsArchived=False AND WorkflowState='Draft' ORDER BY MemberName"
        f!uxAssignReport.Requery
    End If
    f!uxWorkTab0.FontBold = (mode = "Assignments"): f!uxWorkTab1.FontBold = (mode = "Assign"): f!uxWorkTab2.FontBold = (mode = "Contributors")
    UXMyWorkMode = True
End Function
Public Function UXChooseAssignmentReport() As Boolean
    Dim id As Long
    With Forms!frmMyWork
        id = Nz(!uxAssignReport, 0): !txtReportID = id
        If id > 0 Then !txtExpectedRevision = CurrentRevision(id) Else !txtExpectedRevision = 0
        !lblAssignment.Caption = "This assignment will use the currently saved revision of " & Nz(DLookup("MemberName", "Report", "ReportID=" & id), "the selected report") & "."
    End With
    UXChooseAssignmentReport = True
End Function
Public Function UXLoadMyWork() As Boolean
    If Not LoadMyWork() Then Exit Function
    UXMyWorkMode "Assignments"
    UXLoadMyWork = True
End Function
Public Function UXDashboardRefresh() As Boolean
    Dim f As Access.Form, pid As Long, recoveryCount As Long
    If Not CurrentProject.AllForms("frmHome").IsLoaded Then Exit Function
    Set f = Forms!frmHome
    If Not UXHasControl(f, "uxRecentReports") Then Exit Function
    pid = Nz(DLookup("ActiveProfileID", "Workspace"), 0)
    f!uxReportCount.Caption = CStr(DCount("*", "Report", "IsArchived=False"))
    f!uxGroupCount.Caption = CStr(DCount("*", "SummaryGroup", "IsArchived=False"))
    f!uxWorkCount.Caption = CStr(DCount("*", "ReviewTask", "AssignedTo=" & pid & " AND TaskState='Pending'"))
    recoveryCount = DCount("*", "RecoveryDraft", "AuthorProfileID=" & pid)
    f!uxRecoveryStatus.Caption = IIf(recoveryCount > 0, recoveryCount & " recovery checkpoint(s) available. Review before continuing that report.", "No recovery checkpoints. Save draft creates a saved revision.")
    f!cmdRecovery.Enabled = (recoveryCount > 0)
    f!uxRecentReports.RowSource = "SELECT TOP 5 ReportID,MemberName AS Member,FormKind AS [Report type],PeriodStart AS [Period from],PeriodEnd AS [Period to],UXReportState(ReportID) AS [Local status] FROM Report WHERE IsArchived=False ORDER BY ModifiedAt DESC,ReportID DESC"
    f!uxRecentReports.Requery
    f!uxContext.Caption = Nz(DLookup("CommandName", "Workspace"), "Swift NAVFIT") & " | Contributor: " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & pid), "")
    UXDashboardRefresh = True
End Function
Public Function UXOpenRecentReport() As Boolean
    Dim id As Long
    id = Nz(Forms!frmHome!uxRecentReports, 0): If id = 0 Then Exit Function
    If Not UXLeaveEditor() Then Exit Function
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    UXOpenRecentReport = True
End Function
Public Function UXChooseNewType(ByVal kind As String) As Boolean
    TempVars!UXNewKind = kind
    DoCmd.Close acForm, "frmUXNewReport", acSaveNo
    UXChooseNewType = True
End Function
Public Function UXNewReport() As Boolean
    Dim kind As String
    If Not UXLeaveEditor() Then Exit Function
    On Error Resume Next: TempVars.Remove "UXNewKind": On Error GoTo 0
    TempVars.Add "UXNewKind", ""
    DoCmd.OpenForm "frmUXNewReport", , , , , acDialog
    kind = CStr(TempVars!UXNewKind): TempVars.Remove "UXNewKind"
    If Len(kind) = 0 Then Exit Function
    DoCmd.OpenForm "frmEditor", , , , , , "0"
    Forms!frmEditor!cboKind = kind
    Call UXEditorKindChanged()
    UXNewReport = True
End Function
Public Function UXClearReportFilters() As Boolean
    With Forms!frmLibrary
        !txtSearch = Null: !cboArchive = "Active": !lstFolders = 0: !uxFolderFilter = 0
        !uxTypeFilter = "All": !uxStatusFilter = "All"
    End With
    UXClearReportFilters = RefreshLibrary()
End Function
