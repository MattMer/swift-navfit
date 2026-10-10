Attribute VB_Name = "modUXGroupPages"
Option Explicit
Private createGroup As Long, createGroupStamp As String, createDefaultStamp As String
Private createdReport As Long, createdReportGroup As Long
Public Function UXGroupNavigate(ByVal screen As String, ByVal page As Long) As Boolean
    Dim id As Long, modes As Variant
    On Error GoTo Failed
    id = Nz(Forms(screen)!uxGroupContext, 0)
    If id = 0 Then Exit Function
    If screen = "frmUXComparison" And (page = 0 Or page = 3) Then
        UXGroupTabState screen, page
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
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then DoCmd.Close acForm, "frmUXComparison", acSaveNo
    If Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then DoCmd.OpenForm "frmMajorWorkbench", , , , , acHidden, "ContextOnly"
    Forms!frmMajorWorkbench!cboGroup = id
    modes = Array("Members", "Defaults", "Review", "Comparison", "History")
    If (page = 0 Or page = 3) And DCount("*", "GroupMember", "GroupID=" & id) > 0 Then
        Forms!frmMajorWorkbench!uxMode = CStr(modes(page))
    Else
        UXWorkbenchMode CStr(modes(page))
        If page <> 1 Then
            Forms!frmMajorWorkbench.Visible = True
            Forms!frmMajorWorkbench.SetFocus
        End If
    End If
    Select Case page
        Case 0, 3
            If DCount("*", "GroupMember", "GroupID=" & id) > 0 Then
                UXOpenComparison
                Forms!frmUXComparison!uxGroupContext = id
                UXGroupTabState "frmUXComparison", page
            End If
        Case 1
            UXOpenGroupDefaults
            If CurrentProject.AllForms("frmUXBulk").IsLoaded Then Forms!frmUXBulk!uxGroupContext = id
    End Select
    UXGroupNavigate = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Summary group"
End Function
Public Function UXGroupTabState(ByVal screen As String, ByVal page As Long) As Boolean
    Dim i As Long, f As Access.Form
    Set f = Forms(screen)
    For i = 0 To 4
        f.Controls("uxGroupTab" & i).FontBold = (i = page)
        f.Controls("uxGroupTab" & i).Visible = (Nz(f!uxGroupContext, 0) > 0)
    Next
    If UXHasControl(f, "uxAddGroupReport") Then f!uxAddGroupReport.Visible = (Nz(f!uxGroupContext, 0) > 0)
    UXGroupTabState = True
End Function
Public Function UXGroupAddReport(ByVal screen As String) As Boolean
    Dim f As Access.Form
    On Error GoTo Failed
    If CurrentProject.AllForms("frmUXBulk").IsLoaded Then
        If UXGroupBulkPending() Then
            Forms!frmUXBulk!uxStatus.Caption = "Apply the previewed defaults or close this tool before adding a report. New reports use saved starting defaults."
            Forms!frmUXBulk.SetFocus: Exit Function
        End If
    End If
    createGroup = Nz(Forms(screen)!uxGroupContext, 0)
    If createGroup = 0 Then Exit Function
    If Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & createGroup), True) Then Err.Raise vbObjectError + 7061, , "Recover this group before adding reports."
    createGroupStamp = GroupStamp(createGroup): createDefaultStamp = UXDefaultsStamp(createGroup)
    DoCmd.OpenForm "frmUXGroupReport"
    Set f = Forms!frmUXGroupReport
    f!lblGroup.Caption = DLookup("Title", "SummaryGroup", "GroupID=" & createGroup) & " | " & DLookup("FormKind", "SummaryGroup", "GroupID=" & createGroup) & " | " & DLookup("Paygrade", "SummaryGroup", "GroupID=" & createGroup)
    f!cboMember.RowSource = "SELECT MemberID,FullName,RankRate,ReportName,MemberSSN,VersionCounter FROM CommandMember WHERE IsArchived=False ORDER BY FullName"
    f!txtFrom = UXEffectiveDefault(createGroup, "PeriodStart"): f!txtTo = UXEffectiveDefault(createGroup, "PeriodEnd")
    UXGroupAddReport = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Add report"
End Function
Public Function UXGroupReportMember() As Boolean
    Dim f As Access.Form
    Set f = Forms!frmUXGroupReport
    If IsNull(f!cboMember) Then Exit Function
    f!txtName = f!cboMember.Column(1): f!txtRank = f!cboMember.Column(2): f!txtOfficial = f!cboMember.Column(3): f!txtSSN = f!cboMember.Column(4): f!txtMemberVersion = f!cboMember.Column(5)
    UXGroupReportMember = True
End Function
Public Function CreateGroupDraft(ByVal groupID As Long, ByVal groupBaseline As String, ByVal defaultBaseline As String, ByVal memberID As Long, ByVal memberVersion As Long, ByVal name As String, ByVal official As String, ByVal rank As String, ByVal ssn As String, ByVal startEntry As Variant, ByVal endEntry As Variant) As Long
    Dim w As DAO.Workspace, values As Object, key As Variant, kind As String, grade As String, id As Long, startDate As Variant, endDate As Variant, failure As Long, message As String
    If Trim$(name) = "" Or Len(name) > 120 Then Err.Raise vbObjectError + 7061, , "Enter a member name within 120 characters."
    If Len(official) > 27 Or Len(rank) > 5 Then Err.Raise vbObjectError + 7061, , "Report name permits 27 characters; rank/rate permits five."
    If ssn <> "" Then
        If Len(ssn) <> 9 Or ssn Like "*[!0-9]*" Then Err.Raise vbObjectError + 7061, , "Enter nine SSN digits or leave SSN unentered."
    End If
    startDate = Null: endDate = Null
    If Len(Nz(startEntry, "")) > 0 Then
        If Not IsDate(startEntry) Then Err.Raise vbObjectError + 7061, , "Enter a valid From date."
        startDate = DateValue(startEntry)
    End If
    If Len(Nz(endEntry, "")) > 0 Then
        If Not IsDate(endEntry) Then Err.Raise vbObjectError + 7061, , "Enter a valid To date."
        endDate = DateValue(endEntry)
    End If
    If Not IsNull(startDate) And Not IsNull(endDate) Then
        If endDate < startDate Then Err.Raise vbObjectError + 7061, , "To date must follow From date."
    End If
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If GroupStamp(groupID) <> groupBaseline Or UXDefaultsStamp(groupID) <> defaultBaseline Then Err.Raise vbObjectError + 7061, , "Group or defaults changed. Close and reopen Add report."
    If Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & groupID), True) Then Err.Raise vbObjectError + 7061, , "Recover this group before adding reports."
    kind = DLookup("FormKind", "SummaryGroup", "GroupID=" & groupID): grade = DLookup("Paygrade", "SummaryGroup", "GroupID=" & groupID)
    If ValidationWornPaygrade(rank) <> grade Then Err.Raise vbObjectError + 7061, , "Enter a rank/rate matching group paygrade " & grade & "."
    If memberID > 0 Then
        If Nz(DLookup("IsArchived", "CommandMember", "MemberID=" & memberID), True) Or Nz(DLookup("VersionCounter", "CommandMember", "MemberID=" & memberID), 0) <> memberVersion Then Err.Raise vbObjectError + 7061, , "Member profile changed. Select the member again."
        If ssn <> Nz(DLookup("MemberSSN", "CommandMember", "MemberID=" & memberID), "") Then Err.Raise vbObjectError + 7061, , "Edit SSN in the member profile first, then select the member again."
        If Not IsNull(startDate) And Not IsNull(endDate) Then
            If DCount("*", "Report", "ReportID In (SELECT ReportID FROM MemberReport WHERE MemberID=" & memberID & ") AND FormKind=" & QuoteSQL(kind) & " AND PeriodStart=#" & Format$(startDate, "yyyy-mm-dd") & "# AND PeriodEnd=#" & Format$(endDate, "yyyy-mm-dd") & "#") > 0 Then Err.Raise vbObjectError + 7061, , "A report already exists for this member, type and period. Open that report instead."
        End If
    End If
    Set values = CreateObject("Scripting.Dictionary")
    For Each key In DraftFieldKeys(): values.Add CStr(key), Null: Next
    For Each key In UXSharedKeys(): values(CStr(key)) = UXEffectiveDefault(groupID, CStr(key)): Next
    values("OfficialName") = official: values("RankRate") = rank: values("MemberSSN") = IIf(ssn = "", Null, ssn)
    DirectoryCreationContext memberID
    id = CommitDraft(0, kind, kind & " - " & name, name, startDate, endDate, "", "Create from summary group", values("OfficialComments"), _
        FieldsSubset(values, AdministrationKeys()), FieldsSubset(values, TraitKeys()), FieldsSubset(values, AssessmentKeys()), values("PromotionStatus"), FieldsSubset(values, ReportingSeniorKeys()), _
        FieldsSubset(values, ReportContextKeys()), FieldsSubset(values, MemberDetailsKeys()), FieldsSubset(values, ServiceDetailsKeys()), FieldsSubset(values, DecisionDetailsKeys()), FieldsSubset(values, ReportPeopleKeys()), FieldsSubset(values, PartyNamesKeys()), values("CommandEmployment"))
    DirectoryCreationContext 0: AddGroupMember groupID, id, grade
    w.CommitTrans: CreateGroupDraft = id: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next: DirectoryCreationContext 0: w.Rollback
    On Error GoTo 0: Err.Raise failure, "CreateGroupDraft", message
End Function
Public Function UXGroupReportCreate() As Boolean
    Dim f As Access.Form, id As Long
    On Error GoTo Failed
    Set f = Forms!frmUXGroupReport
    id = CreateGroupDraft(createGroup, createGroupStamp, createDefaultStamp, Nz(f!cboMember, 0), Nz(f!txtMemberVersion, 0), Nz(f!txtName, ""), Nz(f!txtOfficial, ""), Nz(f!txtRank, ""), Nz(f!txtSSN, ""), f!txtFrom, f!txtTo)
    DoCmd.Close acForm, f.Name, acSaveNo
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then UXCloseComparison
    If CurrentProject.AllForms("frmUXBulk").IsLoaded Then UXCloseBlocks
    UXWorkbenchMode "Members"
    createdReport = id: createdReportGroup = createGroup
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    Forms!frmEditor!cmdClose.Caption = "Back to group members"
    Forms!frmEditor!cmdClose.OnClick = "=UXGroupCreatedReportClose()"
    UXGroupReportCreate = True: Exit Function
Failed:
    If id > 0 Then
        MsgBox "The report was saved and added to the group. Open it from the member list. " & Err.Description, vbExclamation, "Open created report"
    Else
        f!lblStatus.Caption = Err.Description
    End If
End Function
Public Function UXGroupCreatedReportClose() As Boolean
    Dim id As Long, groupID As Long
    On Error GoTo Failed
    id = Nz(Forms!frmEditor!txtReportID, 0): groupID = createdReportGroup
    If Not UXLeaveEditor() Then Exit Function
    If id = createdReport And groupID > 0 Then
        If Not CurrentProject.AllForms("frmMajorWorkbench").IsLoaded Then OpenScreen "frmMajorWorkbench"
        Forms!frmMajorWorkbench!cboGroup = groupID: UXWorkbenchMode "Members"
        If Not Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & groupID), True) Then UXGroupNavigate "frmMajorWorkbench", 0
    End If
    createdReport = 0: createdReportGroup = 0: UXGroupCreatedReportClose = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Return to summary group"
End Function
