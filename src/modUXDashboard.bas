Attribute VB_Name = "modUXDashboard"
Option Compare Binary
Option Explicit
Public Function UXIdentity(ByVal official As Variant, ByVal localName As Variant) As String
    UXIdentity = Trim$(Nz(official, ""))
    If Len(UXIdentity) = 0 Then UXIdentity = Trim$(Nz(localName, ""))
    If Len(UXIdentity) = 0 Then UXIdentity = "Unnamed member"
End Function
Public Function UXMemberLabel(ByVal reportID As Long) As String
    UXMemberLabel = UXIdentity(DLookup("OfficialName", "Report", "ReportID=" & reportID), DLookup("MemberName", "Report", "ReportID=" & reportID))
End Function
Public Function UXRevisionMemberLabel(ByVal revisionID As Long) As String
    UXRevisionMemberLabel = UXIdentity(MajorRevisionValue(revisionID, "OfficialName"), MajorRevisionValue(revisionID, "MemberName"))
End Function
Public Function UXReportState(ByVal reportID As Long) As String
    Dim revision As Long, routes As DAO.Recordset
    If Nz(DLookup("IsArchived", "Report", "ReportID=" & reportID), False) Then UXReportState = "Archived": Exit Function
    If Nz(DLookup("WorkflowState", "Report", "ReportID=" & reportID), "Draft") = "Approved" Then UXReportState = "Local approval": Exit Function
    revision = Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & reportID), 0)
    If DCount("*", "ReviewTask", "ReportID=" & reportID & " AND TaskState='Pending' AND RevisionID<>" & revision) > 0 Then UXReportState = "Review needs restart": Exit Function
    If DCount("*", "RouteRun", "ReportID=" & reportID & " AND RunState In ('Pending','Accepted') AND RevisionID<>" & revision) > 0 Then UXReportState = "Review needs restart": Exit Function
    Set routes = CurrentDb.OpenRecordset("SELECT RunID FROM RouteRun WHERE ReportID=" & reportID & " AND RunState In ('Pending','Accepted')", dbOpenSnapshot)
    Do While Not routes.EOF
        If Not RouteIsCurrent(routes!RunID) Then routes.Close: UXReportState = "Review needs restart": Exit Function
        routes.MoveNext
    Loop
    routes.Close
    If DCount("*", "ReviewTask", "ReportID=" & reportID & " AND TaskState='Pending'") > 0 Or DCount("*", "RouteRun", "ReportID=" & reportID & " AND RunState='Pending'") > 0 Then UXReportState = "In review": Exit Function
    If DCount("*", "ReviewTask", "ReportID=" & reportID & " AND TaskState='Returned' AND RevisionID=" & revision) > 0 Then UXReportState = "Returned": Exit Function
    UXReportState = "Draft"
End Function
Public Function UXGroupAverage(ByVal groupID As Long) As Variant
    Dim db As DAO.Database, members As DAO.Recordset, traits As DAO.Recordset, values As Object, key As Variant, token As String, value As Variant, total As Double, marks As Long
    Set db = CurrentDb: Set values = CreateObject("Scripting.Dictionary")
    Set members = db.OpenRecordset("SELECT r.CurrentRevisionID,r.ObservationStatus FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & groupID & " AND r.IsArchived=False", dbOpenSnapshot)
    Set traits = db.OpenRecordset("SELECT v.RevisionID,v.FieldKey,v.ValueText,v.IsNullValue FROM (RevisionValue AS v INNER JOIN Report AS r ON v.RevisionID=r.CurrentRevisionID) INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & groupID & " AND r.IsArchived=False AND v.FieldKey In ('Trait33','Trait34','Trait35','Trait36','Trait37','Trait38','Trait39')", dbOpenSnapshot)
    Do While Not traits.EOF
        token = traits!RevisionID & ":" & traits!FieldKey
        If traits!IsNullValue Then values(token) = Null Else values(token) = traits!ValueText.Value
        traits.MoveNext
    Loop
    traits.Close
    Do While Not members.EOF
        If Nz(members!ObservationStatus, "") <> "NOB" Then
            For Each key In TraitKeys()
                token = Nz(members!CurrentRevisionID, 0) & ":" & key
                If Not values.Exists(token) Then Err.Raise vbObjectError + 4007, "modUXDashboard.UXGroupAverage", "Snapshot field missing."
                value = values(token)
                Select Case Nz(value, "")
                    Case "1.0", "2.0", "3.0", "4.0", "5.0": total = total + CDbl(value): marks = marks + 1
                End Select
            Next
        End If
        members.MoveNext
    Loop
    members.Close
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
    UXShowGroupMembers id
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
    UXRestoreViewport f
    If Not UXHasControl(f, "uxGroupTitle") Then Exit Function
    id = Nz(f!cboGroup, 0)
    If mode = "Members" And id > 0 Then UXFocusedGroupMode = UXShowGroupMembers(id, True): Exit Function
    If id = 0 Then f.ScrollBars = 2: f.Section(acDetail).Height = 12500: Exit Function
    f.ScrollBars = 0: f.Section(acDetail).Height = 8600
    f!uxGroupTitle.Caption = Nz(DLookup("Title", "SummaryGroup", "GroupID=" & id), "Summary group")
    average = UXGroupAverage(id): Set psg = RSGroupPSG(id)
    f!uxGroupMeta.Caption = Nz(DLookup("FormKind", "SummaryGroup", "GroupID=" & id), "") & " | " & Nz(DLookup("Paygrade", "SummaryGroup", "GroupID=" & id), "") & " | " & DCount("*", "GroupMember", "GroupID=" & id) & " members | Trait average: " & IIf(IsNull(average), "Unavailable", Format$(Nz(average, 0), "0.00")) & " | PSG RSCA: " & IIf(IsNull(psg("Value")), "Unavailable", Format$(Nz(psg("Value"), 0), "0.00") & IIf(psg("Estimated"), " (estimate)", ""))
    For Each c In f.Controls
        If Left$(c.Tag, 4) = "WBX:" Then c.Visible = False
    Next
    f!lblStatus.Visible = True: f!uxBack.Visible = True
    If UXHasControl(f, "uxValidation") Then f!uxValidation.Visible = True
    Select Case mode
        Case "Members": f!lblStatus.Caption = "Open a member report or edit its saved values in the member comparison."
        Case "Review": f!lblStatus.Caption = "Select a member to review its saved revision. Open the report to see its full content."
        Case "History": f!lblStatus.Caption = "Inspect saved batch values and preview undo before applying it."
    End Select
    If mode = "Review" Then
        GroupPosition "lstReports", 2850, 3000, 10450, 2900
        f!lstReports.RowSource = "SELECT r.ReportID,UXIdentity(r.OfficialName,r.MemberName) AS Member,r.FormKind AS [Report type],r.PeriodEnd AS [Period to],UXReportState(r.ReportID) AS [Local status],UXTableITA(r.ReportID) AS ITA,r.PromotionRecommendation AS [Promotion recommendation] FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & id & " ORDER BY UXIdentity(r.OfficialName,r.MemberName)"
        f!lstReports.ColumnCount = 7: f!lstReports.ColumnWidths = "0;2400;1000;1700;2000;800;2400": f!lstReports.Requery
        f!uxOpenGroupReport.Visible = True
        If mode = "Review" And UXHasControl(f, "uxFocusedReview") Then f!uxFocusedReview.Visible = True
    ElseIf mode = "Comparison" Then
        f!uxComparison.Visible = True
        f!lblStatus.Caption = "Compare saved trait marks and stage proposed edits. A comparison baseline is captured automatically."
    ElseIf mode = "Defaults" Then
        f!cboScope = "Group"
        GroupPosition "uxDefaultsEditor", 2850, 3000, 10450, 550
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
        f.Controls("uxMode" & i).Visible = False
        f.Controls("uxMode" & i).FontBold = (modes(i) = mode)
    Next
    UXResizeWorkspace f.Name
    f!uxGroupContext = id
    UXGroupTabState f.Name, Switch(mode = "Members", 0, mode = "Defaults", 1, mode = "Review", 2, mode = "Comparison", 3, mode = "History", 4), psg
    UXGroupLocks f, id
    UXFocusedGroupMode = True
End Function
Public Function UXComparisonPage(ByVal page As String) As Boolean
    Dim f As Access.Form, c As Access.Control, preview As Boolean
    Set f = Forms!frmUXComparison: UXRestoreViewport f: preview = (page = "Preview")
    f!uxTraitGrid.Visible = Not preview: f!uxGridAverage.Visible = Not preview: f!uxProposalDetail.Visible = preview
    For Each c In f.Controls
        If c.Name Like "cboTrait##" Or c.Name Like "lblTrait##" Or c.Name = "lstTraits" Or c.Name = "uxMember" Or c.Name = "uxITA" Or c.Name = "cboPromotion" Or c.Name = "cmdStage" Then c.Visible = False
    Next
    f!lstProposals.Visible = preview: f!cmdApply.Visible = preview
    For Each c In f.Controls
        If c.Tag = "GroupMemberAction" Then c.Visible = (Not preview And Nz(f!txtMeeting, 0) = 0)
    Next
    f!uxMarksPage.FontBold = Not preview: f!uxPreviewPage.FontBold = preview
    UXGridViewport
    UXComparisonPage = True
End Function
Public Function UXMyWorkMode(ByVal mode As String) As Boolean
    Dim f As Access.Form, c As Access.Control, names As Variant, name As Variant
    Set f = Forms!frmMyWork
    UXRestoreViewport f
    If Not UXHasControl(f, "uxWorkTab0") Then Exit Function
    f!uxWorkTab0.SetFocus
    For Each c In f.Controls
        If Left$(c.Tag, 4) = "MWX:" Then c.Visible = False
    Next
    Select Case mode
        Case "Contributors": names = Array("cboContributor", "cmdSwitch", "txtContributor", "cmdAdd", "lblScope")
        Case "Notes": names = Array("lblReportNotes", "lstReportNotes", "cmdOpenReportNote", "cmdRefreshReportNotes")
        Case "Assign": names = Array("uxAssignReport", "uxAssignCaption", "lblAssignment", "lblReviewer", "cboReviewer", "lblDue", "txtDue", "lblNotes", "txtNotes", "cmdAssign")
        Case Else: names = Array("chkAll", "lblAll", "cmdRefresh", "lstWork", "uxWorkStatus", "cmdOpenWork", "txtHistory", "lblNotes", "txtNotes", "cmdAccept", "cmdReturn", "cmdCancel", "cmdDraft")
    End Select
    For Each name In names: f.Controls(CStr(name)).Visible = True: Next
    f!lblNotes.Caption = IIf(mode = "Assign", "Assignment instructions", "Decision note (required when returning a report)")
    If mode = "Assign" Then
        f!uxAssignReport.RowSource = "SELECT ReportID,UXIdentity(OfficialName,MemberName) & ' | ' & FormKind & ' | ' & Nz(Format(PeriodEnd,'yyyy-mm-dd'),'No period') AS Description FROM Report WHERE IsArchived=False AND WorkflowState='Draft' ORDER BY UXIdentity(OfficialName,MemberName)"
        f!uxAssignReport.Requery
    End If
    f!uxWorkNotesTab.FontBold = (mode = "Notes")
    f!uxWorkTab0.FontBold = (mode = "Assignments"): f!uxWorkTab1.FontBold = (mode = "Assign"): f!uxWorkTab2.FontBold = (mode = "Contributors")
    UXResizeWorkspace f.Name
    UXMyWorkMode = True
    If mode = "Notes" Then UXMyWorkNotes
    If mode = "Assignments" Then Call UXLoadWorkQueue
    Call UXWorkLayout
End Function
Public Function UXChooseAssignmentReport() As Boolean
    Dim id As Long
    With Forms!frmMyWork
        id = Nz(!uxAssignReport, 0): !txtReportID = id
        If id > 0 Then !txtExpectedRevision = CurrentRevision(id) Else !txtExpectedRevision = 0
        !lblAssignment.Caption = "This assignment will use the currently saved revision of " & UXMemberLabel(id) & "."
    End With
    UXChooseAssignmentReport = True
End Function
Public Function UXLoadMyWork() As Boolean
    ' The Assignments mode owns the initial unified queue refresh.
    If Not LoadMyWork(False) Then Exit Function
    UXMyWorkMode "Assignments"
    UXLoadMyWork = True
End Function
Public Function UXDashboardRefresh() As Boolean
    Dim f As Access.Form, pid As Long, recoveryCount As Long
    If Not CurrentProject.AllForms("frmHome").IsLoaded Then Exit Function
    Set f = Forms!frmHome
    If Not UXHasControl(f, "uxRecentReports") Then Exit Function
    pid = ActiveContributorID()
    If UXHasControl(f, "uxContributor") Then f!uxContributor = pid
    f!uxReportCount.Caption = CStr(DCount("*", "Report", "IsArchived=False"))
    f!uxGroupCount.Caption = CStr(DCount("*", "SummaryGroup", "IsArchived=False"))
    f!uxWorkCount.Caption = CStr(UXWorkCount())
    recoveryCount = DCount("*", "RecoveryDraft", "AuthorProfileID=" & pid)
    If UXHasControl(f, "uxNoteRecovery") Then
        f!uxNoteRecovery.Caption = "Unposted notes (" & DCount("*", "ReportNoteDraft", "AuthorProfileID=" & pid) & ")"
        f!uxNoteRecovery.Enabled = DCount("*", "ReportNoteDraft", "AuthorProfileID=" & pid) > 0
    End If
    f!uxRecoveryStatus.Caption = IIf(recoveryCount > 0, recoveryCount & " recovery checkpoint(s) available. Review before continuing that report.", "No recovery checkpoints. Save draft creates a saved revision.")
    f!cmdRecovery.Enabled = (recoveryCount > 0)
    f!uxRecentReports.RowSource = "SELECT TOP 5 ReportID,UXIdentity(OfficialName,MemberName) AS Member,FormKind AS [Report type],PeriodStart AS [Period from],PeriodEnd AS [Period to],UXReportState(ReportID) AS [Local status] FROM Report WHERE IsArchived=False ORDER BY ModifiedAt DESC,ReportID DESC"
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
