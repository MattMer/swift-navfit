Attribute VB_Name = "modRSDirectoryUI"
Option Compare Binary
Option Explicit
Private seniorInitial As Object
Private seniorInitialTitle As String

Public Function RSLeaveSeniorEditor() As Boolean
    Dim f As Access.Form, decision As String, pendingBaseline As Boolean
    RSLeaveSeniorEditor = True
    If Not CurrentProject.AllForms("frmUXSeniorEditor").IsLoaded Then Exit Function
    Set f = Forms!frmUXSeniorEditor
    pendingBaseline = Len(Nz(f!txtAsOf, "")) > 0 Or Len(Nz(f!txtRSCA, "")) > 0 Or Len(Nz(f!txtTraitCount, "")) > 0 Or Len(Nz(f!txtTraitSum, "")) > 0 Or Len(Nz(f!txtSource, "")) > 0
    If RSSeniorDirty() Or pendingBaseline Then
        decision = UXAsk("Senior")
        If decision = "Cancel" Then RSLeaveSeniorEditor = False: Exit Function
        If decision = "Save" Then
            If Not RSSaveSeniorUI() Then RSLeaveSeniorEditor = False: Exit Function
            If pendingBaseline Then
                If Not RSAddBaselineUI() Then RSLeaveSeniorEditor = False: Exit Function
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
Public Function RSLoadDirectory() As Boolean
    Forms!frmUXSeniors!lstSeniors.RowSource = "SELECT SeniorID,Title,VersionCounter FROM ReportingSeniorRecord ORDER BY Title"
    Forms!frmUXSeniors!lstSeniors.Requery: RSLoadDirectory = True
End Function
Public Function RSOpenSenior(Optional ByVal createNew As Boolean = False) As Boolean
    Dim id As Long
    If Not createNew Then id = Nz(Forms!frmUXSeniors!lstSeniors, 0)
    If Not createNew And id = 0 Then Forms!frmUXSeniors!lblStatus.Caption = "Choose a reporting senior record.": Exit Function
    If CurrentProject.AllForms("frmUXSeniorEditor").IsLoaded Then Forms!frmUXSeniorEditor.SetFocus: Exit Function
    DoCmd.OpenForm "frmUXSeniorEditor", , , , , , CStr(id): RSOpenSenior = True
End Function
Public Function RSLoadSeniorEditor() As Boolean
    Dim f As Access.Form, key As Variant, id As Long
    Set f = Forms!frmUXSeniorEditor: id = CLng(Nz(f.OpenArgs, "0")): f!txtSeniorID = id
    Set seniorInitial = RSSeniorValues(id)
    seniorInitialTitle = Nz(MajorScalar("SELECT Title FROM ReportingSeniorRecord WHERE SeniorID=" & id), "")
    f!txtTitle = seniorInitialTitle: f!txtVersion = Nz(MajorScalar("SELECT VersionCounter FROM ReportingSeniorRecord WHERE SeniorID=" & id), 0)
    For Each key In RSSeniorKeys(): f.Controls("rs" & key) = seniorInitial(CStr(key)): Next
    RSReloadBaselines: RSSelectSeniorPage 0: RSLoadSeniorEditor = True
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
    id = RSSaveSenior(Nz(f!txtSeniorID, 0), Nz(f!txtTitle, ""), values, Nz(f!txtVersion, 0))
    f!txtSeniorID = id: f!txtVersion = MajorScalar("SELECT VersionCounter FROM ReportingSeniorRecord WHERE SeniorID=" & id)
    Set seniorInitial = values: seniorInitialTitle = Nz(f!txtTitle, "")
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
    RSSelectSeniorPage = True
End Function
Public Function RSReloadBaselines() As Boolean
    Forms!frmUXSeniorEditor!lstBaselines.RowSource = "SELECT Paygrade,AsOfDate,CurrentRSCA,TraitCount,TraitSum,SourceNote FROM RSCABaseline WHERE SeniorID=" & Nz(Forms!frmUXSeniorEditor!txtSeniorID, 0) & " ORDER BY AsOfDate DESC,BaselineID DESC"
    Forms!frmUXSeniorEditor!lstBaselines.Requery: RSReloadBaselines = True
End Function
Public Function RSAddBaselineUI() As Boolean
    Dim f As Access.Form, count As Double, average As Variant, total As Variant, dateEntry As String, asOf As Date, id As Long
    On Error GoTo Failed
    Set f = Forms!frmUXSeniorEditor
    If Nz(f!txtSeniorID, 0) = 0 Or RSSeniorDirty() Then Err.Raise vbObjectError + 4820, , "Save the reporting senior details before adding a historical baseline."
    If Len(Nz(f!txtTraitCount, "")) = 0 Or Not IsNumeric(f!txtTraitCount) Then Err.Raise vbObjectError + 4821, , "Enter a whole previous trait count."
    count = CDbl(f!txtTraitCount)
    If count <> Fix(count) Or count < 0 Or count > 2147483647# Then Err.Raise vbObjectError + 4821, , "Enter a whole trait count from 0 to 2147483647."
    average = f!txtRSCA.Value: If Len(Nz(average, "")) = 0 Then average = Null
    total = f!txtTraitSum.Value: If Len(Nz(total, "")) = 0 Then total = Null
    dateEntry = Nz(UXDateEntry(f!txtAsOf.Value), "")
    If Not dateEntry Like "####-##-##" Then Err.Raise vbObjectError + 4822, , "Enter an as-of date as YYYY-MM-DD."
    asOf = DateSerial(CInt(Left$(dateEntry, 4)), CInt(Mid$(dateEntry, 6, 2)), CInt(Right$(dateEntry, 2)))
    If Format$(asOf, "yyyy-mm-dd") <> dateEntry Then Err.Raise vbObjectError + 4822, , "Enter a valid as-of date."
    id = RSAddBaseline(f!txtSeniorID, Nz(f!cboPaygrade, ""), asOf, average, CLng(count), total, Nz(f!txtSource, ""))
    RSReloadBaselines: f!lblStatus.Caption = "Historical baseline saved. Counts are graded traits, not reports. Exact trait sum avoids rounded-RSCA estimates."
    f!txtRSCA = Null: f!txtTraitCount = Null: f!txtTraitSum = Null: f!txtSource = Null: f!txtAsOf = Null
    RSAddBaselineUI = True: Exit Function
Failed:
    Forms!frmUXSeniorEditor!lblStatus.Caption = Err.Description & " Pending baseline entries are retained."
End Function
Public Function RSUseSeniorInEditor() As Boolean
    Dim values As Object, key As Variant, id As Long
    On Error GoTo Failed
    id = Nz(Forms!frmEditor!uxSeniorRecord, 0): If id = 0 Then Err.Raise vbObjectError + 4800, , "Choose a reporting senior record."
    If Not Forms!frmEditor!cmdSave.Enabled Then Err.Raise vbObjectError + 4823, , "Reopen the report before editing."
    Set values = RSSeniorValues(id)
    For Each key In RSSeniorKeys(): Forms!frmEditor.Controls("txt" & key) = values(CStr(key)): Next
    UXRefreshEditor: RSUseSeniorInEditor = True: Exit Function
Failed:
    Forms!frmEditor!uxError.Caption = Err.Description
End Function
Public Function RSOpenGroupCalculator() As Boolean
    Dim id As Long
    id = Nz(Forms!frmMajorWorkbench!cboGroup, 0): If id = 0 Then Exit Function
    DoCmd.OpenForm "frmUXGroupRSCA", , , , , , CStr(id): RSOpenGroupCalculator = True
End Function
Public Function RSLoadGroupCalculator() As Boolean
    Dim f As Access.Form, id As Long
    Set f = Forms!frmUXGroupRSCA: id = CLng(f.OpenArgs): f!txtGroupID = id
    f!lblHeading.Caption = Nz(MajorScalar("SELECT Title FROM SummaryGroup WHERE GroupID=" & id), "Group PSG RSCA")
    f!cboSenior.RowSource = "SELECT SeniorID,Title FROM ReportingSeniorRecord ORDER BY Title"
    f!cboSenior = MajorScalar("SELECT SeniorID FROM GroupRSCABaseline WHERE GroupID=" & id)
    RSGroupSeniorChanged
    f!cboBaseline = MajorScalar("SELECT BaselineID FROM GroupRSCABaseline WHERE GroupID=" & id)
    f!chkExcludes = Nz(MajorScalar("SELECT ExcludesGroup FROM GroupRSCABaseline WHERE GroupID=" & id), False)
    Call RSShowGroupCalculation: RSLoadGroupCalculator = True
End Function
Public Function RSGroupSeniorChanged() As Boolean
    Dim f As Access.Form, grade As String
    Set f = Forms!frmUXGroupRSCA: grade = Nz(MajorScalar("SELECT Paygrade FROM SummaryGroup WHERE GroupID=" & f!txtGroupID), "")
    f!cboBaseline = Null: f!chkExcludes = False
    f!cboBaseline.RowSource = "SELECT BaselineID,Format(AsOfDate,'yyyy-mm-dd') & ' / RSCA ' & Nz(CurrentRSCA,'none') & ' / ' & TraitCount & ' traits' FROM RSCABaseline WHERE SeniorID=" & Nz(f!cboSenior, 0) & " AND Paygrade=" & QuoteSQL(grade) & " ORDER BY AsOfDate DESC,BaselineID DESC"
    f!cboBaseline.Requery: f!lblStatus.Caption = "Choose a historical baseline for " & grade & ", confirm it excludes this group, then save the selection."
    RSGroupSeniorChanged = True
End Function
Public Function RSSaveGroupCalculation() As Boolean
    On Error GoTo Failed
    With Forms!frmUXGroupRSCA
        RSSetGroupBaseline CLng(!txtGroupID), Nz(!cboSenior, 0), Nz(!cboBaseline, 0), Nz(!chkExcludes, False)
    End With
    Call RSShowGroupCalculation: RSSaveGroupCalculation = True: Exit Function
Failed:
    Forms!frmUXGroupRSCA!lblStatus.Caption = Err.Description
End Function
Public Function RSShowGroupCalculation() As Boolean
    Dim result As Object
    Set result = RSGroupPSG(CLng(Forms!frmUXGroupRSCA!txtGroupID))
    With Forms!frmUXGroupRSCA
        If IsNull(result("Value")) Then
            !lblCalculation.Caption = "PSG RSCA: unavailable"
        Else
            !lblCalculation.Caption = "PSG RSCA (saved group): " & Format$(result("Value"), "0.00") & IIf(result("Estimated"), " - estimate", "")
        End If
        !txtCalculationContext = UXGroupCalculationText(CLng(!txtGroupID))
        !lblStatus.Caption = result("Reason") & IIf(result("Partial"), ". Some traits are blank/NOB; only graded numeric traits are included.", "") & IIf(result("Estimated"), ". Enter the exact historical trait sum to calculate from exact totals.", "")
    End With
    RSShowGroupCalculation = True
End Function
Public Function RSStageGroupSenior() As Boolean
    Dim values As Object, key As Variant
    On Error GoTo Failed
    Forms!frmMajorWorkbench!cboGroup = Forms!frmUXGroupRSCA!txtGroupID
    Set values = RSSeniorValues(Nz(Forms!frmUXGroupRSCA!cboSenior, 0))
    If Nz(Forms!frmUXGroupRSCA!cboSenior, 0) = 0 Then Err.Raise vbObjectError + 4800, , "Choose a reporting senior."
    If Not UXStartGroupBulk() Then Exit Function
    For Each key In RSSeniorKeys()
        Forms!frmUXBulk.Controls("uxValue" & key) = values(CStr(key)): UXBulkChanged CStr(key)
    Next
    Forms!frmUXBulk!txtReason = "Apply selected reporting senior details to group members"
    UXBulkSection 1: RSStageGroupSenior = True: Exit Function
Failed:
    Forms!frmUXGroupRSCA!lblStatus.Caption = Err.Description
End Function
