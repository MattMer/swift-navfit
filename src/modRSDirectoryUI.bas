Attribute VB_Name = "modRSDirectoryUI"
Option Compare Binary
Option Explicit
Private loadingGroupCalculation As Boolean
Private seniorInitial As Object
Private seniorInitialTitle As String
Private seniorInitialMember As Variant
Private pickerMode As String
Private pickerRUAD As Object

Public Function RSLeaveSeniorEditor() As Boolean
    Dim f As Access.Form, decision As String, pendingBaseline As Boolean
    RSLeaveSeniorEditor = True
    If Not CurrentProject.AllForms("frmUXSeniorEditor").IsLoaded Then Exit Function
    Set f = Forms!frmUXSeniorEditor
    pendingBaseline = RSTableDirty()
    If RSSeniorDirty() Or pendingBaseline Then
        decision = UXAsk("Senior")
        If decision = "Cancel" Then RSLeaveSeniorEditor = False: Exit Function
        If decision = "Save" Then
            If Not RSSaveSeniorUI() Then RSLeaveSeniorEditor = False: Exit Function
            If pendingBaseline Then
                If Not RSSaveTable() Then RSLeaveSeniorEditor = False: Exit Function
            End If
        End If
    End If
    DoCmd.Close acForm, "frmUXSeniorEditor", acSaveNo
    Set seniorInitial = Nothing
End Function

Public Function RSOpenDirectory() As Boolean
    If Not UXLeaveEditor() Then Exit Function
    DoCmd.OpenForm "frmUXSeniors": RSOpenDirectory = True
End Function
Private Function SeniorValueSQL(ByVal key As String, ByVal alias As String) As String
    SeniorValueSQL = "(SELECT First(v.ValueText) FROM ReportingSeniorValue AS v WHERE v.SeniorID=s.SeniorID AND v.FieldKey='" & key & "') AS [" & alias & "]"
End Function
Public Function RSLoadDirectory() As Boolean
    Forms!frmUXSeniors!lstSeniors.RowSource = "SELECT s.SeniorID,s.Title," & SeniorValueSQL("ReportingSeniorGrade", "SeniorGrade") & "," & SeniorValueSQL("ReportingSeniorDesignator", "SeniorDesig") & "," & _
        SeniorValueSQL("ReportingSeniorTitle", "SeniorTitle") & "," & SeniorValueSQL("ReportingSeniorUIC", "SeniorUIC") & ",IIf(s.MemberID Is Null,'','Linked') AS DirectoryLink FROM ReportingSeniorRecord AS s ORDER BY s.Title"
    Forms!frmUXSeniors!lstSeniors.Requery: RSLoadDirectory = True
End Function
'New records start from the member picker; opening a record goes straight to it.
Public Function RSOpenSenior(Optional ByVal createNew As Boolean = False) As Boolean
    Dim id As Long
    If Not createNew Then id = Nz(Forms!frmUXSeniors!lstSeniors, 0)
    If Not createNew And id = 0 Then Forms!frmUXSeniors!lblStatus.Caption = "Choose a reporting senior record.": Exit Function
    If CurrentProject.AllForms("frmUXSeniorEditor").IsLoaded Then Forms!frmUXSeniorEditor.SetFocus: Exit Function
    If createNew Then
        If CurrentProject.AllForms("frmUXSeniorPicker").IsLoaded Then DoCmd.Close acForm, "frmUXSeniorPicker", acSaveNo
        DoCmd.OpenForm "frmUXSeniorPicker", , , , , , "new": RSOpenSenior = True: Exit Function
    End If
    DoCmd.OpenForm "frmUXSeniorEditor", , , , , , CStr(id): RSOpenSenior = True
End Function
'OpenArgs is a record ID, or "new:<MemberID>" for a new record filled from the directory.
Public Function RSLoadSeniorEditor() As Boolean
    Dim f As Access.Form, key As Variant, id As Long, args As String, member As Long
    Set f = Forms!frmUXSeniorEditor: args = Nz(f.OpenArgs, "0")
    If Left$(args, 4) = "new:" Then member = CLng(Mid$(args, 5)) Else id = CLng(args)
    f!txtSeniorID = id
    Set seniorInitial = RSSeniorValues(id)
    seniorInitialTitle = Nz(MajorScalar("SELECT Title FROM ReportingSeniorRecord WHERE SeniorID=" & id), "")
    seniorInitialMember = MajorScalar("SELECT MemberID FROM ReportingSeniorRecord WHERE SeniorID=" & id)
    f!txtTitle = seniorInitialTitle: f!txtVersion = Nz(MajorScalar("SELECT VersionCounter FROM ReportingSeniorRecord WHERE SeniorID=" & id), 0)
    f!txtMemberID = seniorInitialMember
    For Each key In RSSeniorKeys(): f.Controls("rs" & key) = seniorInitial(CStr(key)): Next
    Call RSTableReset: RSReloadBaselines: RSSelectSeniorPage 0
    If member > 0 Then
        f!txtTitle = Nz(MajorScalar("SELECT IIf(Len(Nz(ReportName,""""))=0,FullName,ReportName) FROM CommandMember WHERE MemberID=" & member), "")
        RSApplyMemberLink member
    Else
        RSShowLink
    End If
    RSLoadSeniorEditor = True
End Function
'Control values without reading the active control's text, which needs focus.
Private Function SeniorControlValues() As Object
    Dim result As Object, key As Variant
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In RSSeniorKeys(): result.Add CStr(key), Forms!frmUXSeniorEditor.Controls("rs" & key).Value: Next
    Set SeniorControlValues = result
End Function
Private Function PendingMember() As Variant
    PendingMember = Null
    If Nz(Forms!frmUXSeniorEditor!txtMemberID, 0) <> 0 Then PendingMember = CLng(Forms!frmUXSeniorEditor!txtMemberID)
End Function
Public Function RSShowLink() As Boolean
    Dim f As Access.Form, member As Variant, text As String, differs As String
    Set f = Forms!frmUXSeniorEditor: member = PendingMember()
    If IsNull(member) Then
        f!lblDirectoryLink.Caption = "Not linked to a command member. Link one to fill name, grade, designator, UIC and SSN from the member directory."
        f!cmdLinkMember.Caption = "Link to member"
    Else
        text = Nz(MajorScalar("SELECT IIf(Len(Nz(ReportName,""""))=0,FullName,ReportName) & IIf(Len(Nz(RankRate,""""))=0,'',' (' & RankRate & ')') FROM CommandMember WHERE MemberID=" & member), "a missing member")
        differs = RSMemberDifferences(member, SeniorControlValues())
        f!lblDirectoryLink.Caption = "Directory member: " & text & IIf(Len(differs) > 0, ". Directory differs: " & differs & ".", ". Matches the directory.")
        f!cmdLinkMember.Caption = "Change member"
    End If
    RSShowLink = True
End Function
'Link the editor to a member and copy the values the profile has (pending until saved).
Public Function RSApplyMemberLink(ByVal memberID As Long) As Boolean
    Dim f As Access.Form, values As Object, key As Variant
    Set f = Forms!frmUXSeniorEditor: Set values = RSMemberValues(memberID)
    f!txtMemberID = memberID
    For Each key In RSMemberKeys()
        If Not IsNull(values(CStr(key))) Then f.Controls("rs" & key) = values(CStr(key))
    Next
    RSShowLink
    f!lblStatus.Caption = "Directory details copied. Check the title, then save record details. Saved reports remain unchanged."
    RSApplyMemberLink = True
End Function
Public Function RSLinkMember() As Boolean
    If CurrentProject.AllForms("frmUXSeniorPicker").IsLoaded Then DoCmd.Close acForm, "frmUXSeniorPicker", acSaveNo
    DoCmd.OpenForm "frmUXSeniorPicker", , , , , , "link": RSLinkMember = True
End Function
Public Function RSLinkUpdate() As Boolean
    Dim member As Variant
    member = PendingMember()
    If IsNull(member) Then Forms!frmUXSeniorEditor!lblStatus.Caption = "Link a command member first.": Exit Function
    RSLinkUpdate = RSApplyMemberLink(CLng(member))
End Function
Public Function RSUnlinkMember() As Boolean
    Forms!frmUXSeniorEditor!txtMemberID = Null: RSShowLink
    Forms!frmUXSeniorEditor!lblStatus.Caption = "Directory link removed. Typed values stay. Save record details to keep this change."
    RSUnlinkMember = True
End Function
'Member picker. Designator and UIC come from each member's RUAD profile.
Public Function RSMemberRUAD(ByVal memberID As Long, ByVal key As String) As Variant
    If pickerRUAD Is Nothing Then Set pickerRUAD = CreateObject("Scripting.Dictionary")
    If Not pickerRUAD.Exists(memberID) Then pickerRUAD.Add memberID, RUADDecode(MajorScalar("SELECT RUADData FROM CommandMember WHERE MemberID=" & memberID))
    RSMemberRUAD = pickerRUAD(memberID)(key)
End Function
Public Function RSLoadPicker() As Boolean
    Dim f As Access.Form
    Set f = Forms!frmUXSeniorPicker: Set pickerRUAD = Nothing
    pickerMode = Nz(f.OpenArgs, "new")
    If pickerMode = "link" Then
        f.Caption = "Link reporting senior to a member": f!lblHeading.Caption = f.Caption
        f!cmdManual.Visible = False
    End If
    RSLoadPicker = RSPickerFilter(): UXSortHeaders f
End Function
Private Function LikeText(ByVal text As String) As String
    LikeText = Replace(Replace(Replace(Replace(text, "[", "[[]"), "*", "[*]"), "?", "[?]"), "#", "[#]")
End Function
Public Function RSPickerFilter() As Boolean
    Dim f As Access.Form, search As String, where As String
    Set f = Forms!frmUXSeniorPicker: search = Trim$(Nz(f!txtSearch, ""))
    where = "m.IsArchived=False"
    If Len(search) > 0 Then
        search = QuoteSQL("*" & LikeText(search) & "*")
        where = where & " AND (m.FullName Like " & search & " OR m.ReportName Like " & search & " OR m.RankRate Like " & search & " OR m.RUADData Like " & search & ")"
    End If
    f!lstMembers.RowSource = "SELECT m.MemberID,IIf(Len(Nz(m.ReportName,''))=0,m.FullName,m.ReportName) AS MemberName,m.RankRate,RSMemberRUAD(m.MemberID,'DESG') AS MemberDesig,RSMemberRUAD(m.MemberID,'UIC') AS MemberUIC,IIf(s.SeniorID Is Null,'','Already a senior') AS SeniorStatus " & _
        "FROM CommandMember AS m LEFT JOIN ReportingSeniorRecord AS s ON s.MemberID=m.MemberID WHERE " & where & " ORDER BY IIf(Len(Nz(m.ReportName,''))=0,m.FullName,m.ReportName)"
    f!lstMembers.Requery
    f!lblStatus.Caption = IIf(f!lstMembers.ListCount = 0, "No active members match. Change the search, or enter details by hand.", "Choose a member, then Use selected member.")
    RSPickerFilter = True
End Function
Public Function RSPickerUse() As Boolean
    Dim f As Access.Form, member As Long, owner As Long, current As Long
    Set f = Forms!frmUXSeniorPicker: member = Nz(f!lstMembers, 0)
    If member = 0 Then f!lblStatus.Caption = "Choose a member first.": Exit Function
    If pickerMode = "link" Then
        If Not CurrentProject.AllForms("frmUXSeniorEditor").IsLoaded Then f!lblStatus.Caption = "The reporting senior record is closed. Open it again to link a member.": Exit Function
        current = Nz(Forms!frmUXSeniorEditor!txtSeniorID, 0)
    End If
    owner = Nz(MajorScalar("SELECT SeniorID FROM ReportingSeniorRecord WHERE MemberID=" & member & " AND SeniorID<>" & current), 0)
    If owner > 0 Then f!lblStatus.Caption = "This member is already the reporting senior record """ & Nz(MajorScalar("SELECT Title FROM ReportingSeniorRecord WHERE SeniorID=" & owner), "") & """. Open that record instead.": Exit Function
    DoCmd.Close acForm, "frmUXSeniorPicker", acSaveNo
    If pickerMode = "link" Then
        RSPickerUse = RSApplyMemberLink(member)
    Else
        DoCmd.OpenForm "frmUXSeniorEditor", , , , , , "new:" & member: RSPickerUse = True
    End If
End Function
Public Function RSPickerManual() As Boolean
    DoCmd.Close acForm, "frmUXSeniorPicker", acSaveNo
    DoCmd.OpenForm "frmUXSeniorEditor", , , , , , "0": RSPickerManual = True
End Function
Public Function RSReadSeniorUI() As Object
    Dim result As Object, key As Variant, c As Access.Control
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In RSSeniorKeys()
        Set c = Forms!frmUXSeniorEditor.Controls("rs" & key)
        If Screen.ActiveControl.Name = c.Name Then result.Add CStr(key), c.Text Else result.Add CStr(key), c.Value
    Next
    Set RSReadSeniorUI = result
End Function
Public Function RSSeniorDirty() As Boolean
    Dim values As Object, key As Variant
    If seniorInitial Is Nothing Then Exit Function
    Set values = RSReadSeniorUI()
    If Nz(Forms!frmUXSeniorEditor!txtTitle, "") <> seniorInitialTitle Then RSSeniorDirty = True: Exit Function
    If Nz(PendingMember(), 0) <> Nz(seniorInitialMember, 0) Then RSSeniorDirty = True: Exit Function
    For Each key In RSSeniorKeys()
        If Not SameFieldValue(values(CStr(key)), seniorInitial(CStr(key))) Then RSSeniorDirty = True: Exit Function
    Next
End Function
Public Function RSPendingLabel() As Boolean
    Forms!frmUXSeniorEditor!lblStatus.Caption = "Unsaved record changes. Saved reports remain unchanged.": RSPendingLabel = True
End Function
Public Function RSSaveSeniorUI() As Boolean
    Dim f As Access.Form, values As Object, id As Long
    On Error GoTo Failed
    Set f = Forms!frmUXSeniorEditor: Set values = RSReadSeniorUI()
    id = RSSaveSenior(Nz(f!txtSeniorID, 0), Nz(f!txtTitle, ""), values, Nz(f!txtVersion, 0), PendingMember())
    f!txtSeniorID = id: f!txtVersion = MajorScalar("SELECT VersionCounter FROM ReportingSeniorRecord WHERE SeniorID=" & id)
    Set seniorInitial = values: seniorInitialTitle = Nz(f!txtTitle, ""): seniorInitialMember = PendingMember()
    RSShowLink
    f!lblStatus.Caption = "Record saved. Saved report values remain unchanged."
    If CurrentProject.AllForms("frmUXSeniors").IsLoaded Then RSLoadDirectory
    RSSaveSeniorUI = True: Exit Function
Failed:
    Forms!frmUXSeniorEditor!lblStatus.Caption = "Unable to save: " & Err.Description & " Pending values are retained."
End Function
Public Function RSSelectSeniorPage(ByVal page As Long) As Boolean
    Dim c As Access.Control
    Forms!frmUXSeniorEditor!cmdSave.SetFocus
    For Each c In Forms!frmUXSeniorEditor.Controls
        If Left$(c.Tag, 3) = "SR:" Then c.Visible = (CLng(Mid$(c.Tag, 4)) = page)
    Next
    RSTablePage page
    RSSelectSeniorPage = True
End Function
Public Function RSReloadBaselines() As Boolean
    Call RSShowHistory: RSTableReloadVisible: RSReloadBaselines = True
End Function
Public Function RSAddBaselineUI() As Boolean
    Dim f As Access.Form, count As Double, average As Variant, total As Variant, dateEntry As String, asOf As Date, id As Long
    On Error GoTo Failed
    Set f = Forms!frmUXSeniorEditor
    If Nz(f!txtSeniorID, 0) = 0 Or RSSeniorDirty() Then Err.Raise vbObjectError + 4820, "modRSDirectoryUI.RSAddBaselineUI", "Save the reporting senior details before adding a historical baseline."
    If Len(Nz(f!txtTraitCount, "")) = 0 Or Not IsNumeric(f!txtTraitCount) Then Err.Raise vbObjectError + 4821, "modRSDirectoryUI.RSAddBaselineUI", "Enter a whole previous trait count."
    count = CDbl(f!txtTraitCount)
    If count <> Fix(count) Or count < 0 Or count > 2147483647# Then Err.Raise vbObjectError + 4821, "modRSDirectoryUI.RSAddBaselineUI", "Enter a whole trait count from 0 to 2147483647."
    average = f!txtRSCA.Value: If Len(Nz(average, "")) = 0 Then average = Null
    total = f!txtTraitSum.Value: If Len(Nz(total, "")) = 0 Then total = Null
    dateEntry = Nz(UXDateEntry(f!txtAsOf.Value), "")
    If Not dateEntry Like "####-##-##" Then Err.Raise vbObjectError + 4822, "modRSDirectoryUI.RSAddBaselineUI", "Enter an as-of date as YYYY-MM-DD."
    asOf = DateSerial(CInt(Left$(dateEntry, 4)), CInt(Mid$(dateEntry, 6, 2)), CInt(Right$(dateEntry, 2)))
    If Format$(asOf, "yyyy-mm-dd") <> dateEntry Then Err.Raise vbObjectError + 4822, "modRSDirectoryUI.RSAddBaselineUI", "Enter a valid as-of date."
    id = RSAddBaseline(f!txtSeniorID, Nz(f!cboPaygrade, ""), asOf, average, CLng(count), total, Nz(f!txtSource, ""))
    Call RSReloadBaselines: f!lblStatus.Caption = "Historical baseline saved. Counts are graded traits, not reports. Exact trait sum avoids rounded-RSCA estimates."
    f!txtRSCA = Null: f!txtTraitCount = Null: f!txtTraitSum = Null: f!txtSource = Null: f!txtAsOf = Null
    RSAddBaselineUI = True: Exit Function
Failed:
    Forms!frmUXSeniorEditor!lblStatus.Caption = Err.Description & " Pending baseline entries are retained."
End Function
Public Function RSUseSeniorInEditor() As Boolean
    Dim values As Object, key As Variant, id As Long
    On Error GoTo Failed
    id = Nz(Forms!frmEditor!uxSeniorRecord, 0): If id = 0 Then Err.Raise vbObjectError + 4800, "modRSDirectoryUI.RSUseSeniorInEditor", "Choose a reporting senior record."
    If Not Forms!frmEditor!cmdSave.Enabled Then Err.Raise vbObjectError + 4823, "modRSDirectoryUI.RSUseSeniorInEditor", "Reopen the report before editing."
    Set values = RSSeniorValues(id)
    For Each key In RSSeniorKeys(): Forms!frmEditor.Controls("txt" & key) = values(CStr(key)): Next
    Call UXRefreshEditor: RSUseSeniorInEditor = True: Exit Function
Failed:
    Forms!frmEditor!uxError.Caption = Err.Description
End Function
Public Function RSOpenGroupCalculator(Optional ByVal groupID As Long = 0) As Boolean
    Dim id As Long
    id = groupID: If id = 0 Then id = Nz(Forms!frmMajorWorkbench!cboGroup, 0)
    If id = 0 Then Exit Function
    If CurrentProject.AllForms("frmUXGroupRSCA").IsLoaded Then DoCmd.Close acForm, "frmUXGroupRSCA", acSaveNo
    DoCmd.OpenForm "frmUXGroupRSCA", , , , , , CStr(id): RSOpenGroupCalculator = True
End Function
Public Function RSLoadGroupCalculator() As Boolean
    Dim f As Access.Form, id As Long
    On Error GoTo Failed
    loadingGroupCalculation = True
    Set f = Forms!frmUXGroupRSCA: id = CLng(f.OpenArgs): f!txtGroupID = id
    f!lblHeading.Caption = Nz(MajorScalar("SELECT Title FROM SummaryGroup WHERE GroupID=" & id), "Group PSG RSCA")
    f!cboSenior.RowSource = "SELECT SeniorID,Title FROM ReportingSeniorRecord ORDER BY Title"
    f!cboSenior = MajorScalar("SELECT SeniorID FROM GroupRSCABaseline WHERE GroupID=" & id)
    RSGroupSeniorChanged
    f!cboBaseline = MajorScalar("SELECT BaselineID FROM GroupRSCABaseline WHERE GroupID=" & id)
    f!chkExcludes = Nz(MajorScalar("SELECT ExcludesGroup FROM GroupRSCABaseline WHERE GroupID=" & id), False)
    loadingGroupCalculation = False
    Call RSShowGroupCalculation: RSLoadGroupCalculator = True
    Exit Function
Failed:
    loadingGroupCalculation = False
    Err.Raise Err.Number, "RSLoadGroupCalculator", Err.Description
End Function
Public Function RSGroupSeniorChanged() As Boolean
    Dim f As Access.Form, grade As String
    Set f = Forms!frmUXGroupRSCA: grade = Nz(MajorScalar("SELECT Paygrade FROM SummaryGroup WHERE GroupID=" & f!txtGroupID), "")
    f!cboBaseline = Null: f!chkExcludes = False
    f!cboBaseline.RowSource = "SELECT BaselineID,Format(AsOfDate,'yyyy-mm-dd') & ' / RSCA ' & Nz(CurrentRSCA,'none') & ' / ' & TraitCount & ' traits' FROM RSCABaseline WHERE SeniorID=" & Nz(f!cboSenior, 0) & " AND Paygrade=" & QuoteSQL(grade) & " ORDER BY AsOfDate DESC,BaselineID DESC"
    f!cboBaseline.Requery: f!lblStatus.Caption = "Choose historical totals for " & grade & ". Selections save automatically. Confirm these reports are not included in the BOL totals."
    If Not loadingGroupCalculation Then RSSaveGroupCalculation
    RSGroupSeniorChanged = True
End Function
Public Function RSGroupBaselineChanged() As Boolean
    Forms!frmUXGroupRSCA!chkExcludes = False
    RSGroupBaselineChanged = RSSaveGroupCalculation()
End Function
Public Function RSSaveGroupCalculation() As Boolean
    On Error GoTo Failed
    With Forms!frmUXGroupRSCA
        RSSetGroupBaseline CLng(!txtGroupID), Nz(!cboSenior, 0), Nz(!cboBaseline, 0), Nz(!chkExcludes, False)
    End With
    Call RSShowGroupCalculation: UXRefreshGroupSenior: RSSaveGroupCalculation = True: Exit Function
Failed:
    Forms!frmUXGroupRSCA!lblStatus.Caption = Err.Description
End Function
Public Function RSShowGroupCalculation() As Boolean
    Dim result As Object, loaded As DAO.Recordset
    Set loaded = RSGroupRows(CLng(Forms!frmUXGroupRSCA!txtGroupID))
    Set result = RSGroupPSG(CLng(Forms!frmUXGroupRSCA!txtGroupID), loaded)
    With Forms!frmUXGroupRSCA
        If IsNull(result("Value")) Then
            !lblCalculation.Caption = "PSG RSCA: unavailable"
        Else
            !lblCalculation.Caption = "PSG RSCA (saved group): " & Format$(result("Value"), "0.00") & IIf(result("Estimated"), " - estimate", "")
        End If
        !txtCalculationContext = UXGroupCalculationText(CLng(!txtGroupID), loaded)
        !lblStatus.Caption = result("Reason") & IIf(result("Partial"), ". Some traits are blank/NOB; only graded numeric traits are included.", "") & IIf(result("Estimated"), ". Enter the exact historical trait sum to calculate from exact totals.", "")
    End With
    loaded.Close
    RSShowGroupCalculation = True
End Function
Public Function RSStageGroupSenior(Optional ByVal decision As VbMsgBoxResult = 0) As Boolean
    Dim values As Object, groupID As Long
    On Error GoTo Failed
    groupID = Nz(Forms!frmUXGroupRSCA!txtGroupID, 0)
    If Nz(Forms!frmUXGroupRSCA!cboSenior, 0) = 0 Then Err.Raise vbObjectError + 4800, "modRSDirectoryUI.RSStageGroupSenior", "Choose a reporting senior."
    Set values = RSSeniorValues(Nz(Forms!frmUXGroupRSCA!cboSenior, 0))
    RSStageGroupSenior = UXConfirmReportFields(groupID, values, "Apply selected reporting senior details", decision)
    If RSStageGroupSenior Then
        Call RSShowGroupCalculation: UXRefreshGroupSenior
        Forms!frmUXGroupRSCA!lblStatus.Caption = "Reporting senior details saved to group members. " & Forms!frmUXGroupRSCA!lblStatus.Caption
    End If
    Exit Function
Failed:
    Forms!frmUXGroupRSCA!lblStatus.Caption = Err.Description & " No partial changes were saved."
End Function
