Attribute VB_Name = "modUXGroupPages"
Option Explicit
Private createGroup As Long, createGroupStamp As String, createDefaultStamp As String
Private createdReport As Long, createdReportGroup As Long
'One live Members surface for new, empty and populated groups.
Public Function UXShowGroupMembers(ByVal id As Long, Optional ByVal refresh As Boolean = False) As Boolean
    Dim gridLoaded As Boolean
    If id <= 0 Then Exit Function
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then
        If Not UXGridDrain() Then Exit Function
        If Nz(Forms!frmUXComparison!uxGroupContext, 0) = id And Nz(Forms!frmUXComparison!txtMeeting, 0) = 0 Then
            If refresh Then UXGridLoad: gridLoaded = True
        Else
            If Not UXCloseComparison(False) Then Exit Function
        End If
    End If
    If CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then Forms!frmMajorWorkbench.Visible = False
    If Not CurrentProject.AllForms("frmUXComparison").IsLoaded Then DoCmd.OpenForm "frmUXComparison", , , , , , "G:" & id: gridLoaded = True
    Forms!frmUXComparison.Caption = Nz(MajorScalar("SELECT Title FROM SummaryGroup WHERE GroupID=" & id), "Summary group")
    Forms!frmUXComparison!lblHeading.Caption = Forms!frmUXComparison.Caption
    If UXHasControl(Forms!frmUXComparison, "uxGroupMeta") Then Forms!frmUXComparison!uxGroupMeta.Caption = Nz(MajorScalar("SELECT FormKind & ' | ' & Paygrade FROM SummaryGroup WHERE GroupID=" & id), "") & " | " & PersistenceCount("GroupMember", "GroupID=" & id) & " members"
    If Forms!frmUXComparison!uxTraitGrid.Form.Recordset.EOF Then Forms!frmUXComparison!lblStatus.Caption = "No reports in this summary group. Choose group members or add a report."
    'A grid load just set the senior caption through its averages; skip a second PSG pass.
    UXGroupTabState "frmUXComparison", 0, , gridLoaded
    UXGroupLocks Forms!frmUXComparison, id
    Forms!frmUXComparison.SetFocus
    UXShowGroupMembers = True
End Function

'The Members toolbar resolves its own group; legacy tools receive an explicit
'hidden context only when needed, never a visible alternate Members page.
Public Function UXMemberAction(ByVal action As String) As Boolean
    Dim id As Long
    On Error GoTo Failed
    If Not UXGridDrain() Then Exit Function
    id = UXResolvedGroupContext("frmUXComparison"): If id = 0 Then Exit Function
    If action = "Open" Then UXMemberAction = UXGridOpenReport(): Exit Function
    If action = "Validation" Then UXMemberAction = UXOpenValidation(True, id): Exit Function
    If action = "Blocks" Then UXMemberAction = UXStartGroupBulk(False, False, id): Exit Function
    If action = "SignedGroup" Then UXMemberAction = UXSignoffOpenGroup(id, False): Exit Function
    If action = "Signed" Then
        If Forms!frmUXComparison!uxTraitGrid.Form.Recordset.EOF Then Err.Raise vbObjectError + 9633, "modUXGroupPages.UXMemberAction", "Select a member report first."
        UXMemberAction = UXSignoffOpen(CLng(Forms!frmUXComparison!uxTraitGrid.Form!ReportID)): Exit Function
    End If
    If action <> "Members" And action <> "Policy" Then Err.Raise 5, "modUXGroupPages.UXMemberAction", "Unknown member action."
    If Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then DoCmd.OpenForm "frmMajorWorkbench", , , , , acHidden, "ContextOnly"
    Forms!frmMajorWorkbench.Visible = False
    Forms!frmMajorWorkbench!cboGroup = id: Forms!frmMajorWorkbench!uxGroupContext = id
    If action = "Members" Then UXMemberAction = UXChooseMembers() Else UXMemberAction = UXOpenGroupPolicy()
    Exit Function
Failed:
    Forms!frmUXComparison!lblStatus.Caption = Err.Description
End Function
' Route group tabs with one resolved context. Members always uses the live grid;
' only Review/History need the retained workbench adapter. Preserve pending guards.
Public Function UXGroupNavigate(ByVal screen As String, ByVal page As Long) As Boolean
    Dim id As Long, modes As Variant
    On Error GoTo Failed
    If page = 3 Then page = 0 'Legacy Comparison links now resolve to Members.
    If page < 0 Or page > 4 Then Exit Function
    id = UXResolvedGroupContext(screen)
    If id = 0 Then Exit Function
    If screen = "frmUXComparison" And (page = 0 Or page = 3) Then
        UXGroupTabState screen, page
        DoCmd.Echo True
    UXGroupNavigate = True: Exit Function
    End If
    If screen = "frmUXBulk" And page = 1 Then
        If Forms(screen)!lblHeading.Caption = "Group defaults" Then UXGroupNavigate = True: Exit Function
    End If
    If CurrentProject.AllForms("frmUXBulk").IsLoaded Then
        If UXGroupBulkPending() Then
            Forms!frmUXBulk!uxStatus.Caption = "Apply your previewed changes or close this tool to discard them before changing group pages."
            Forms!frmUXBulk.SetFocus: Exit Function
        End If
        DoCmd.Close acForm, "frmUXBulk", acSaveNo
    End If
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then
        If Not UXCloseComparison(False) Then Exit Function
    End If
    If (page = 0 Or page = 1 Or page = 3) And CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then Forms!frmMajorWorkbench.Visible = False
    DoCmd.Echo False
    If page = 0 Or page = 3 Then
        If Not UXShowGroupMembers(id) Then DoCmd.Echo True: Exit Function
    ElseIf page = 1 Then
        If Not UXOpenGroupDefaults(id) Then DoCmd.Echo True: Exit Function
    Else
        If Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then DoCmd.OpenForm "frmMajorWorkbench", , , , , acHidden, "ContextOnly"
        Forms!frmMajorWorkbench.Visible = False
        Forms!frmMajorWorkbench!cboGroup = id
        modes = Array("Members", "Defaults", "Review", "Comparison", "History")
        UXWorkbenchMode CStr(modes(page))
        Forms!frmMajorWorkbench.Visible = True: Forms!frmMajorWorkbench.SetFocus
    End If
    DoCmd.Echo True
    UXGroupNavigate = True: Exit Function
Failed:
    DoCmd.Echo True
    MsgBox Err.Description, vbExclamation, "Summary group"
End Function
' Mark the active group tab and show group actions only when a group context exists.
Public Function UXGroupTabState(ByVal screen As String, ByVal page As Long, Optional ByVal preparedPSG As Object = Nothing, Optional ByVal seniorCaptionCurrent As Boolean = False) As Boolean
    Dim i As Long, f As Access.Form
    Set f = Forms(screen)
    If page = 3 Then page = 0
    For i = 0 To 4
        If UXHasControl(f, "uxGroupTab" & i) Then
            f.Controls("uxGroupTab" & i).FontBold = (i = page)
            f.Controls("uxGroupTab" & i).Visible = (Nz(f!uxGroupContext, 0) > 0 And i <> 3)
        End If
    Next
    If UXHasControl(f, "uxAddGroupReport") Then f!uxAddGroupReport.Visible = (Nz(f!uxGroupContext, 0) > 0)
    If UXHasControl(f, "uxCurrentExport") Then
        f!uxCurrentExport.Visible = (Nz(f!uxGroupContext, 0) > 0)
        f!uxCurrentSenior.Visible = f!uxCurrentExport.Visible: f!uxSeniorCaption.Visible = f!uxCurrentExport.Visible
        If f!uxCurrentExport.Visible And Not seniorCaptionCurrent Then f!uxSeniorCaption.Caption = UXGroupSeniorCaption(CLng(f!uxGroupContext), preparedPSG)
    End If
    UXPositionGroupFooter f
    UXGroupTabState = True
End Function
' Place contextual actions after the visible fields and status of the active page.
Public Sub UXPositionGroupFooter(ByVal f As Access.Form)
    Dim c As Access.Control, bottom As Long
    If Not UXHasControl(f, "uxCurrentExport") Then Exit Sub
    If Not f!uxCurrentExport.Visible Then Exit Sub
    bottom = 2600
    For Each c In f.Controls
        If c.Visible And c.Left >= 200 And c.Tag <> "GroupFooter" And c.Name <> "uxRail" Then
            If c.Top + c.Height > bottom Then bottom = c.Top + c.Height
        End If
    Next
    f!uxSeniorCaption.Top = bottom + 150
    f!uxCurrentSenior.Top = bottom + 500: f!uxCurrentExport.Top = bottom + 500
    If UXHasControl(f, "uxCurrentPDF") Then f!uxCurrentPDF.Top = bottom + 500
    If f.Section(acDetail).Height < bottom + 1100 Then f.Section(acDetail).Height = bottom + 1100
End Sub
' Open report creation from a saved group/default baseline, refusing pending cells,
' unapplied bulk changes or archived groups.
Public Function UXGroupAddReport(ByVal screen As String) As Boolean
    Dim f As Access.Form
    On Error GoTo Failed
    If UXGridPendingCount() > 0 Then Err.Raise vbObjectError + 4814, "modUXGroupPages.UXGroupAddReport", "Retry or discard pending cells before adding a report."
    If CurrentProject.AllForms("frmUXBulk").IsLoaded Then
        If UXGroupBulkPending() Then
            Forms!frmUXBulk!uxStatus.Caption = "Apply the previewed defaults or close this tool before adding a report. New reports use saved starting defaults."
            Forms!frmUXBulk.SetFocus: Exit Function
        End If
    End If
    createGroup = Nz(Forms(screen)!uxGroupContext, 0)
    If createGroup = 0 Then Exit Function
    If Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & createGroup), True) Then Err.Raise vbObjectError + 7061, "modUXGroupPages.UXGroupAddReport", "Recover this group before adding reports."
    createGroupStamp = GroupStamp(createGroup): createDefaultStamp = UXDefaultsStamp(createGroup)
    DoCmd.OpenForm "frmUXGroupReport"
    Set f = Forms!frmUXGroupReport
    f!lblGroup.Caption = DLookup("Title", "SummaryGroup", "GroupID=" & createGroup) & " | " & DLookup("FormKind", "SummaryGroup", "GroupID=" & createGroup) & " | " & DLookup("Paygrade", "SummaryGroup", "GroupID=" & createGroup)
    UXMemberRankChoices f: UXGroupMemberFilter
    f!txtFrom = UXEffectiveDefault(createGroup, "PeriodStart"): f!txtTo = UXEffectiveDefault(createGroup, "PeriodEnd")
    UXGroupAddReport = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Add report"
End Function
Public Function UXGroupMemberFilter() As Boolean
    Dim f As Access.Form
    Set f = Forms!frmUXGroupReport
    f!cboMember = Null
    f!cboMember.RowSource = "SELECT MemberID,FullName,RankRate,ReportName,MemberSSN,VersionCounter FROM CommandMember WHERE IsArchived=False" & UXMemberRankWhere(Nz(f!cboRankFilter, "All ranks/rates")) & " ORDER BY FullName"
    f!cboMember.Requery
    'A changed filter must not retain a hidden member identity or its copied data.
    f!txtName = Null: f!txtOfficial = Null: f!txtRank = Null: f!txtSSN = Null: f!txtMemberVersion = Null
    UXGroupMemberFilter = True
End Function
' Copy the selected directory identity and version into the creation form without saving it.
Public Function UXGroupReportMember() As Boolean
    Dim f As Access.Form
    Set f = Forms!frmUXGroupReport
    If IsNull(f!cboMember) Then Exit Function
    f!txtName = f!cboMember.Column(1): f!txtRank = f!cboMember.Column(2): f!txtOfficial = f!cboMember.Column(3): f!txtSSN = f!cboMember.Column(4): f!txtMemberVersion = f!cboMember.Column(5)
    UXGroupReportMember = True
End Function
' Validate member identity, dates, group/default baselines and duplicate periods, then
' create the report and membership in one transaction using inherited starting values.
Public Function CreateGroupDraft(ByVal groupID As Long, ByVal groupBaseline As String, ByVal defaultBaseline As String, ByVal memberID As Long, ByVal memberVersion As Long, ByVal name As String, ByVal official As String, ByVal rank As String, ByVal ssn As String, ByVal startEntry As Variant, ByVal endEntry As Variant) As Long
    Dim w As DAO.Workspace, values As Object, key As Variant, kind As String, grade As String, id As Long, startDate As Variant, endDate As Variant, failure As Long, message As String
    If Trim$(name) = "" Or Len(name) > 120 Then Err.Raise vbObjectError + 7061, "modUXGroupPages.CreateGroupDraft", "Enter a member name within 120 characters."
    If Len(official) > 27 Or Len(rank) > 5 Then Err.Raise vbObjectError + 7061, "modUXGroupPages.CreateGroupDraft", "Report name permits 27 characters; rank/rate permits five."
    If ssn <> "" Then
        If Len(ssn) <> 9 Or ssn Like "*[!0-9]*" Then Err.Raise vbObjectError + 7061, "modUXGroupPages.CreateGroupDraft", "Enter nine SSN digits or leave SSN unentered."
    End If
    startDate = Null: endDate = Null
    If Len(Nz(startEntry, "")) > 0 Then
        If Not IsDate(startEntry) Then Err.Raise vbObjectError + 7061, "modUXGroupPages.CreateGroupDraft", "Enter a valid From date."
        startDate = DateValue(startEntry)
    End If
    If Len(Nz(endEntry, "")) > 0 Then
        If Not IsDate(endEntry) Then Err.Raise vbObjectError + 7061, "modUXGroupPages.CreateGroupDraft", "Enter a valid To date."
        endDate = DateValue(endEntry)
    End If
    If Not IsNull(startDate) And Not IsNull(endDate) Then
        If endDate < startDate Then Err.Raise vbObjectError + 7061, "modUXGroupPages.CreateGroupDraft", "To date must follow From date."
    End If
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If GroupStamp(groupID) <> groupBaseline Or UXDefaultsStamp(groupID) <> defaultBaseline Then Err.Raise vbObjectError + 7061, "modUXGroupPages.CreateGroupDraft", "Group or defaults changed. Close and reopen Add report."
    If Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & groupID), True) Then Err.Raise vbObjectError + 7061, "modUXGroupPages.CreateGroupDraft", "Recover this group before adding reports."
    kind = DLookup("FormKind", "SummaryGroup", "GroupID=" & groupID): grade = DLookup("Paygrade", "SummaryGroup", "GroupID=" & groupID)
    If ValidationWornPaygrade(rank) <> grade Then Err.Raise vbObjectError + 7061, "modUXGroupPages.CreateGroupDraft", "Enter a rank/rate matching group paygrade " & grade & "."
    If memberID > 0 Then
        If Nz(DLookup("IsArchived", "CommandMember", "MemberID=" & memberID), True) Or Nz(DLookup("VersionCounter", "CommandMember", "MemberID=" & memberID), 0) <> memberVersion Then Err.Raise vbObjectError + 7061, "modUXGroupPages.CreateGroupDraft", "Member profile changed. Select the member again."
        If ssn <> Nz(DLookup("MemberSSN", "CommandMember", "MemberID=" & memberID), "") Then Err.Raise vbObjectError + 7061, "modUXGroupPages.CreateGroupDraft", "Edit SSN in the member profile first, then select the member again."
        If Not IsNull(startDate) And Not IsNull(endDate) Then
            If DCount("*", "Report", "ReportID In (SELECT ReportID FROM MemberReport WHERE MemberID=" & memberID & ") AND FormKind=" & QuoteSQL(kind) & " AND PeriodStart=#" & Format$(startDate, "yyyy-mm-dd") & "# AND PeriodEnd=#" & Format$(endDate, "yyyy-mm-dd") & "#") > 0 Then Err.Raise vbObjectError + 7061, "modUXGroupPages.CreateGroupDraft", "A report already exists for this member, type and period. Open that report instead."
        End If
    End If
    Set values = CreateObject("Scripting.Dictionary")
    For Each key In DraftFieldKeys(): values.Add CStr(key), Null: Next
    For Each key In UXSharedKeys(): values(CStr(key)) = UXEffectiveDefault(groupID, CStr(key)): Next
    values("OfficialName") = official: values("RankRate") = rank: values("MemberSSN") = IIf(ssn = "", Null, ssn)
    If Not IsNull(startDate) And Not IsNull(endDate) Then values("PeriodStart") = startDate: values("PeriodEnd") = endDate
    CounselingCreationDefaults values, memberID
    DirectoryCreationContext memberID
    id = CommitDraft(0, kind, kind & " - " & name, name, startDate, endDate, "", "Create from summary group", values("OfficialComments"), _
        FieldsSubset(values, AdministrationKeys()), FieldsSubset(values, TraitKeys()), FieldsSubset(values, AssessmentKeys()), values("PromotionStatus"), FieldsSubset(values, ReportingSeniorKeys()), _
        FieldsSubset(values, ReportContextKeys()), FieldsSubset(values, MemberDetailsKeys()), FieldsSubset(values, ServiceDetailsKeys()), FieldsSubset(values, DecisionDetailsKeys()), FieldsSubset(values, ReportPeopleKeys()), FieldsSubset(values, PartyNamesKeys()), values("CommandEmployment"))
    DirectoryCreationContext 0: AddGroupMember groupID, id, grade
    w.CommitTrans: CreateGroupDraft = id: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modUXGroupPages", "CreateGroupDraft", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    DirectoryCreationContext 0: SwiftRollback w, swiftPrimary
    On Error GoTo 0: SwiftRaise swiftPrimary
End Function
' Create the pending report from its captured context and open it for editing; retain errors in the dialog.
Public Function UXGroupReportCreate() As Boolean
    Dim f As Access.Form, id As Long
    On Error GoTo Failed
    If UXGridPendingCount() > 0 Then Err.Raise vbObjectError + 4814, "modUXGroupPages.UXGroupReportCreate", "Retry or discard pending cells before creating a report."
    Set f = Forms!frmUXGroupReport
    id = CreateGroupDraft(createGroup, createGroupStamp, createDefaultStamp, Nz(f!cboMember, 0), Nz(f!txtMemberVersion, 0), Nz(f!txtName, ""), Nz(f!txtOfficial, ""), Nz(f!txtRank, ""), Nz(f!txtSSN, ""), f!txtFrom, f!txtTo)
    DoCmd.Close acForm, f.Name, acSaveNo
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then UXCloseComparison
    If CurrentProject.AllForms("frmUXBulk").IsLoaded Then UXCloseBlocks
    If CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then Forms!frmMajorWorkbench.Visible = False
    createdReport = id: createdReportGroup = createGroup
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!cmdClose.Caption = "Back to group members"
    Forms!frmEditor!uxReturnGroup = createGroup
    Forms!frmEditor!cmdClose.OnClick = "=UXGroupCreatedReportClose()"
    UXGroupReportCreate = True: Exit Function
Failed:
    If id > 0 Then
        MsgBox "The report was saved and added to the group. Open it from the member list. " & Err.Description, vbExclamation, "Open created report"
    Else
        f!lblStatus.Caption = Err.Description
    End If
End Function
' Return to the originating group after its newly created report closes.
Public Function UXGroupCreatedReportClose() As Boolean
    Dim id As Long, groupID As Long
    On Error GoTo Failed
    id = Nz(Forms!frmEditor!txtReportID, 0): groupID = Nz(Forms!frmEditor!uxReturnGroup, createdReportGroup)
    If Not UXLeaveEditor() Then Exit Function
    If groupID > 0 Then
        UXShowGroupMembers groupID, True
    End If
    createdReport = 0: createdReportGroup = 0: UXGroupCreatedReportClose = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Return to summary group"
End Function

' One explicit context resolver for group-page transitions; never infer context
' from a different hidden form or an earlier global-defaults session.
Public Function UXResolvedGroupContext(ByVal screen As String) As Long
    Dim f As Access.Form
    Set f = Forms(screen)
    If screen = "frmUXGroups" Then UXResolvedGroupContext = Nz(f!lstGroups, 0)
    If UXHasControl(f, "uxGroupContext") Then UXResolvedGroupContext = Nz(f!uxGroupContext, 0)
    If UXResolvedGroupContext = 0 And screen = "frmMajorWorkbench" Then UXResolvedGroupContext = Nz(f!cboGroup, 0)
End Function
