Attribute VB_Name = "modUX"
Option Compare Binary
Option Explicit
Private currentSection As Long
Private currentPage As Long
Private bulkChanges As Object
Private bulkCSV As String
Private bulkBatch As Long
Private groupContext As Long
Public Function UXFolderFilterChanged() As Boolean
    Forms!frmLibrary!lstFolders = Forms!frmLibrary!uxFolderFilter
    UXFolderFilterChanged = RefreshLibrary()
End Function
Public Function UXHasControl(ByVal f As Access.Form, ByVal name As String) As Boolean
    On Error GoTo Missing
    Dim c As Access.Control: Set c = f.Controls(name): UXHasControl = True
Missing:
End Function
Public Function UXDecide(ByVal decision As String) As Boolean
    TempVars!UXDecision = decision
    DoCmd.Close acForm, "frmUXDecision", acSaveNo
    UXDecide = True
End Function
Public Function UXAsk(ByVal mode As String) As String
    On Error Resume Next: TempVars.Remove "UXDecision": On Error GoTo 0
    TempVars.Add "UXDecision", "Cancel"
    DoCmd.OpenForm "frmUXDecision", , , , , acDialog, mode
    UXAsk = CStr(TempVars!UXDecision): TempVars.Remove "UXDecision"
End Function
Public Function UXLoadDecision() As Boolean
    With Forms!frmUXDecision
        If Nz(.OpenArgs, "") = "Review" Then
            !lblHeading.Caption = "Use which report values?"
            !lblBody.Caption = "Review actions use saved revisions. Save pending changes first, use the saved version while keeping pending edits, or cancel."
            !cmdDiscard.Caption = "Use saved version"
        ElseIf Nz(.OpenArgs, "") = "Senior" Then
            !lblHeading.Caption = "Keep reporting senior changes?"
            !lblBody.Caption = "Save pending record details and any completed RSCA baseline, discard changes, or keep editing."
            !cmdDiscard.Caption = "Discard changes"
        Else
            !lblHeading.Caption = "Keep your report changes?"
            !lblBody.Caption = "Saving creates a report revision. Recovery checkpoints are separate. Discard removes this editor session's pending checkpoint."
            !cmdDiscard.Caption = "Discard changes"
        End If
    End With
    UXLoadDecision = True
End Function
Public Function UXLeaveEditor(Optional ByVal decision As String = "") As Boolean
    If Not CurrentProject.AllForms("frmEditor").IsLoaded Then UXLeaveEditor = True: Exit Function
    If EditorHasPending() Then
        If Len(decision) = 0 Then decision = UXAsk("Leave")
        Select Case decision
            Case "Save": If Not UXSaveEditor() Then Exit Function
            Case "Discard"
            Case Else: Exit Function
        End Select
    End If
    DiscardEditorCheckpoint
    DoCmd.Close acForm, "frmEditor", acSaveNo
    UXLeaveEditor = True
End Function
Public Function UXNavigate(ByVal target As String) As Boolean
    On Error GoTo Failed
    If Not RSLeaveSeniorEditor() Then Exit Function
    If Not UXLeaveEditor() Then Exit Function
    DoCmd.OpenForm target
    UXLoadContext
    UXNavigate = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Unable to open this workspace"
End Function
Public Function UXEditorStart() As Boolean
    currentSection = 0: currentPage = 0
    Call UXSelectSection(0)
    UXEditorStart = True
End Function
Public Function UXSelectSection(ByVal section As Long, Optional ByVal page As Long = 0) As Boolean
    Dim f As Access.Form, c As Access.Control, tags As Variant, id As Long, locked As Boolean, maxPage As Long
    Set f = Forms!frmEditor
    If UXHasControl(f, "cmdClose") Then f!cmdClose.SetFocus
    currentSection = section: currentPage = page: id = Nz(f!txtReportID, 0)
    locked = Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & id), False) Or Nz(MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & id), "Draft") <> "Draft"
    If UXSurfaceAvailable() Then
        For Each c In f.Controls
            If Left$(c.Tag, 6) = "LOCAL:" Then c.Visible = (CLng(Mid$(c.Tag, 7)) = section)
            If c.Name Like "uxSection#" Then
                c.BackColor = IIf(c.Name = "uxSection" & section, 16445671, 16777215)
                c.FontBold = (c.Name = "uxSection" & section)
            End If
        Next
        f!uxSection.Caption = Replace(UXSectionName(section), "&", "&&")
        f!cmdSave.Enabled = Not locked: f!cboKind.Locked = locked
        If locked Then f!uxLocked.Caption = "Read only. Recover archived reports or reopen local approval to edit." Else f!uxLocked.Caption = ""
        If Not UXSurfaceSection(section, locked) Then Exit Function
        UXResizeEditor
        UXRefreshEditor
        UXSelectSection = True: Exit Function
    End If
    For Each c In f.Controls
        If Left$(c.Tag, 3) = "UX:" Then
            tags = Split(c.Tag, ":")
            c.Visible = (CLng(tags(1)) = section And CLng(tags(2)) = page)
            If CLng(tags(1)) = section And CLng(tags(2)) > maxPage Then maxPage = CLng(tags(2))
            If c.ControlType = acTextBox Or c.ControlType = acComboBox Or c.ControlType = acCheckBox Then c.Locked = locked Or tags(3) = "Guidance"
        End If
    Next
    f!uxSection.Caption = Replace(UXSectionName(section), "&", "&&")
    f!uxPage.Caption = "Page " & (page + 1) & " of " & (maxPage + 1)
    f!uxPrevious.Enabled = (page > 0): f!uxNext.Enabled = (page < maxPage)
    f!cmdSave.Enabled = Not locked
    f!cboKind.Locked = locked
    If locked Then f!uxLocked.Caption = "Read only. Reopen local approval or recover the archived report to edit." Else f!uxLocked.Caption = ""
    UXRefreshEditor
    UXSelectSection = True
End Function
Public Function UXEditorKindChanged() As Boolean
    Call EditorTypeChanged()
    UXEditorKindChanged = UXSelectSection(currentSection)
End Function
Public Function UXEditorPage(ByVal delta As Long) As Boolean
    UXEditorPage = UXSelectSection(currentSection, currentPage + delta)
End Function
Public Function UXFieldChanged(ByVal controlName As String, ByVal key As String) As Boolean
    Dim f As Access.Form, value As Variant
    Set f = Forms!frmEditor
    value = f.Controls(controlName).Value
    If key = "ObservationStatus" Then
        If IsNull(value) Then f!cboObservationStatus = Null Else If value Then f!cboObservationStatus = "NOB" Else f!cboObservationStatus = "Observed"
    Else
        If IsNull(value) Then f.Controls("txt" & key) = Null Else If value Then f.Controls("txt" & key) = "Yes" Else f.Controls("txt" & key) = "No"
    End If
    UXRefreshEditor: UXFieldChanged = True
End Function
Public Function UXRefreshEditor() As Boolean
    Dim f As Access.Form, id As Long, r As DAO.Recordset, db As DAO.Database, groups As Long, groupID As Long, total As Double, count As Long, average As Variant, c As Access.Control, key As String, raw As Variant, memberLabel As String
    Set f = Forms!frmEditor
    If Not UXHasControl(f, "uxDirty") Then Exit Function
    f!uxDirty.Caption = IIf(EditorHasPending(), "Unsaved changes", "No pending changes")
    memberLabel = Nz(f!txtMember, "")
    If Len(Trim$(memberLabel)) = 0 Then memberLabel = Nz(f!txtOfficialName, "Unnamed member")
    f!uxReport.Caption = memberLabel & vbCrLf & Nz(f!cboKind, "") & " | " & Nz(f!txtRankRate, "Grade / rate") & " | " & Nz(f!txtStart, "Start date") & " to " & Nz(f!txtEnd, "End date")
    f!uxITA.Caption = "ITA - pending: " & TraitSummary(TraitsFromEditor())
    If Nz(f!cboObservationStatus, "") = "NOB" Then f!uxITA.Caption = "ITA - pending: NOB (whole report)"
    id = Nz(f!txtReportID, 0): Set db = CurrentDb
    f!uxGroup.RowSource = "SELECT g.GroupID,g.Title FROM SummaryGroup AS g INNER JOIN GroupMember AS m ON g.GroupID=m.GroupID WHERE m.ReportID=" & id & " AND g.IsArchived=False ORDER BY g.Title"
    f!uxGroup.Requery
    Set r = db.OpenRecordset("SELECT GroupID FROM GroupMember WHERE ReportID=" & id, dbOpenSnapshot)
    Do While Not r.EOF: groups = groups + 1: groupID = r!GroupID: r.MoveNext: Loop: r.Close
    If groups = 1 And IsNull(f!uxGroup) Then f!uxGroup = groupID
    If groups = 0 Then f!uxGroup = Null
    If Nz(f!uxGroup, 0) > 0 Then
        Set r = db.OpenRecordset("SELECT r.CurrentRevisionID,r.ObservationStatus FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & f!uxGroup & " AND r.IsArchived=False", dbOpenSnapshot)
        Do While Not r.EOF
            If Nz(r!ObservationStatus, "") <> "NOB" Then
                For Each average In TraitKeys()
                    raw = MajorRevisionValue(Nz(r!CurrentRevisionID, 0), CStr(average))
                    Select Case Nz(raw, "")
                        Case "1.0", "2.0", "3.0", "4.0", "5.0": total = total + CDbl(raw): count = count + 1
                    End Select
                Next
            End If
            r.MoveNext
        Loop
        r.Close
        If count > 0 Then f!uxSGA.Caption = "Summary Group Average - saved values: " & Format$(total / count, "0.00") & " (" & count & " graded traits; arithmetic only)" Else f!uxSGA.Caption = "Summary Group Average: unavailable (no graded traits)"
    Else
        f!uxSGA.Caption = "Summary Group Average: " & IIf(groups > 1, "choose the explicit group context", "not assigned to a group")
    End If
    For Each c In f.Controls
        If Left$(c.Name, 5) = "uxChk" Then
            key = Mid$(c.Name, 6)
            If key = "ObservationStatus" Then raw = f!cboObservationStatus.Value Else raw = f.Controls("txt" & key).Value
            If IsNull(raw) Or Len(Nz(raw, "")) = 0 Then c.Value = Null Else If key = "ObservationStatus" Then c.Value = (raw = "NOB") Else c.Value = (raw = "Yes")
        End If
    Next
    If UXSurfaceAvailable() Then UXSurfaceRefresh
    UXRefreshEditor = True
End Function
Public Function UXShowSaveError(ByVal message As String) As Boolean
    Dim f As Access.Form
    Set f = Forms!frmEditor
    If Not UXHasControl(f, "uxError") Then Exit Function
    f!uxError.Caption = "Unable to save: " & message & " Pending values are preserved."
    UXShowSaveError = True
End Function
Public Function UXSaveEditor() As Boolean
    Dim f As Access.Form
    Set f = Forms!frmEditor
    f!uxError.Caption = ""
    If UXSurfaceAvailable() Then
        If Not UXSurfaceCheckAll() Then Exit Function
    End If
    UXSaveEditor = SaveEditor()
    UXRefreshEditor
End Function
Public Function UXLoadGroups() As Boolean
    Forms!frmUXGroups!lstGroups.RowSource = "SELECT g.GroupID,g.Title,g.FormKind,g.Paygrade,Count(m.ReportID) AS Reports FROM SummaryGroup AS g LEFT JOIN GroupMember AS m ON g.GroupID=m.GroupID WHERE g.IsArchived=False GROUP BY g.GroupID,g.Title,g.FormKind,g.Paygrade ORDER BY g.Title"
    Forms!frmUXGroups!lstGroups.Requery
    UXLoadGroups = True
End Function
Public Function UXOpenGroup() As Boolean
    groupContext = Nz(Forms!frmUXGroups!lstGroups, 0)
    If groupContext = 0 Then MsgBox "Choose a summary group.", vbInformation: Exit Function
    DoCmd.OpenForm "frmMajorWorkbench"
    Forms!frmMajorWorkbench!cboGroup = groupContext
    Call UXWorkbenchMode("Members")
    UXOpenGroup = True
End Function
Public Function UXWorkbenchMode(ByVal mode As String) As Boolean
    Dim f As Access.Form, c As Access.Control, tags As Variant
    Set f = Forms!frmMajorWorkbench
    f!uxMode = mode
    For Each c In f.Controls
        If Left$(c.Tag, 4) = "WBX:" Then c.Visible = (InStr(1, ":" & Mid$(c.Tag, 5) & ":", ":" & mode & ":", vbBinaryCompare) > 0)
    Next
    If mode = "Members" Or mode = "Comparison" Or mode = "Defaults" Then f.Caption = Nz(MajorScalar("SELECT Title FROM SummaryGroup WHERE GroupID=" & Nz(f!cboGroup, 0)), "Summary group")
    If mode = "Members" And Nz(f!cboGroup, 0) > 0 Then
        f!lstReports.RowSource = "SELECT r.ReportID,r.Title,r.MemberName,r.FormKind FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & f!cboGroup & " ORDER BY r.MemberName"
        f!lstReports.Requery
    End If
    UXFocusedGroupMode mode
    UXWorkbenchMode = True
End Function
Public Function UXChooseMembers() As Boolean
    Dim f As Access.Form, i As Long, id As Long, kind As String
    groupContext = Nz(Forms!frmMajorWorkbench!cboGroup, 0)
    If groupContext = 0 Then MsgBox "Create or choose a group first.": Exit Function
    kind = Nz(MajorScalar("SELECT FormKind FROM SummaryGroup WHERE GroupID=" & groupContext), "")
    DoCmd.OpenForm "frmUXMembers": Set f = Forms!frmUXMembers
    f!lstReports.RowSource = "SELECT ReportID,MemberName,RankRate,Title FROM Report WHERE IsArchived=False AND FormKind=" & QuoteSQL(kind) & " ORDER BY MemberName"
    f!lstReports.Requery
    For i = 0 To f!lstReports.ListCount - 1
        id = Nz(f!lstReports.Column(0, i), 0)
        f!lstReports.Selected(i) = (DCount("*", "GroupMember", "GroupID=" & groupContext & " AND ReportID=" & id) > 0)
    Next
    UXChooseMembers = True
End Function
Public Function UXSaveMembers() As Boolean
    Dim f As Access.Form, i As Long, id As Long, grade As String, w As DAO.Workspace, started As Boolean
    On Error GoTo Failed
    Set f = Forms!frmUXMembers: grade = Nz(MajorScalar("SELECT Paygrade FROM SummaryGroup WHERE GroupID=" & groupContext), "")
    Set w = DBEngine.Workspaces(0): w.BeginTrans: started = True
    For i = 0 To f!lstReports.ListCount - 1
        id = CLng(f!lstReports.Column(0, i))
        If f!lstReports.Selected(i) Then
            If DCount("*", "GroupMember", "GroupID=" & groupContext & " AND ReportID=" & id) = 0 Then AddGroupMember groupContext, id, grade, Nz(f!txtNote, "")
        Else
            If DCount("*", "GroupMember", "GroupID=" & groupContext & " AND ReportID=" & id) > 0 Then RemoveGroupMember groupContext, id
        End If
    Next
    w.CommitTrans: started = False: DoCmd.Close acForm, "frmUXMembers", acSaveNo
    Call UXWorkbenchMode("Members"): UXSaveMembers = True: Exit Function
Failed:
    If started Then w.Rollback
    Forms!frmUXMembers!lblStatus.Caption = Err.Description
End Function
Public Function UXOpenComparison() As Boolean
    Dim f As Access.Form, id As Long, snapshot As Long, title As String
    On Error GoTo Failed
    Set f = Forms!frmMajorWorkbench
    If CurrentProject.AllForms("frmUXComparison").IsLoaded Then
        Forms!frmUXComparison.SetFocus
        f!lblStatus.Caption = "A comparison is already open. Finish or close it before opening another group comparison."
        Exit Function
    End If
    id = 0
    If id = 0 Then
        snapshot = FreezeSummaryGroup(Nz(f!cboGroup, 0), "Automatically captured for comparison; membership eligibility is not certified.")
        title = Nz(MajorScalar("SELECT Title FROM SummaryGroup WHERE GroupID=" & Nz(f!cboGroup, 0)), "Comparison")
        id = CreateComparisonSet(snapshot, title & " comparison"): f!cboMeeting = id
    End If
    DoCmd.OpenForm "frmUXComparison", , , , , , CStr(id)
    UXOpenComparison = True: Exit Function
Failed:
    Forms!frmMajorWorkbench!lblStatus.Caption = Err.Description
End Function
Public Function UXLoadComparison() As Boolean
    Dim f As Access.Form, block As Long, sql As String
    Set f = Forms!frmUXComparison: f!txtMeeting = CLng(f.OpenArgs)
    f!lblHeading.Caption = Nz(MajorScalar("SELECT Title FROM ComparisonSet WHERE SetID=" & f!txtMeeting), "Comparison")
    sql = "SELECT ReportID,MajorRevisionValue(RevisionID,'MemberName') AS Member"
    For block = 33 To 39: sql = sql & ",MajorRevisionValue(RevisionID,'Trait" & block & "') AS T" & block: Next
    sql = sql & ",NumericTraitAverage(RevisionID) AS ITA,MajorRevisionValue(RevisionID,'PromotionRecommendation') AS Recommendation,RevisionID FROM ComparisonMember WHERE SetID=" & f!txtMeeting & " ORDER BY SortOrdinal"
    f!lstTraits.RowSource = sql: f!lstTraits.Requery
    f!lstProposals.RowSource = "SELECT p.ProposalID,r.MemberName,UXFieldLabel(p.FieldKey),p.ValueText,p.ProposalState FROM MeetingProposal AS p INNER JOIN Report AS r ON p.ReportID=r.ReportID WHERE p.SetID=" & f!txtMeeting & " ORDER BY p.ProposalID DESC"
    f!lstProposals.Requery
    If UXHasControl(f, "uxMarksPage") Then UXComparisonPage "Marks"
    UXLoadComparison = True
End Function
Public Function UXSelectComparisonMember() As Boolean
    Dim f As Access.Form, block As Long, id As Long, locked As Boolean
    Set f = Forms!frmUXComparison: id = Nz(f!lstTraits, 0): If id = 0 Then Exit Function
    For block = 33 To 39: f.Controls("cboTrait" & block) = f!lstTraits.Column(block - 31): Next
    f!cboPromotion = f!lstTraits.Column(10)
    locked = Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & id), True) Or Nz(MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & id), "") <> "Draft"
    f!cmdStage.Enabled = Not locked
    For block = 33 To 39: f.Controls("cboTrait" & block).Locked = locked: Next
    f!cboPromotion.Locked = locked
    f!uxMember.Caption = Nz(f!lstTraits.Column(1), "") & IIf(locked, " - locked; reopen in the report workspace to edit", " - proposed edits; stage and review before Apply")
    UXComparisonAverage: UXSelectComparisonMember = True
End Function
Public Function UXComparisonAverage() As Boolean
    Dim values As Variant, i As Long
    values = Array(Null, Null, Null, Null, Null, Null, Null)
    For i = 0 To 6: values(i) = Forms!frmUXComparison.Controls("cboTrait" & (33 + i)): Next
    Forms!frmUXComparison!uxITA.Caption = "Proposed ITA: " & TraitSummary(values)
    UXComparisonAverage = True
End Function
Public Function UXStageComparison() As Boolean
    Dim f As Access.Form, id As Long, revision As Long, block As Long, before As Variant, value As Variant, count As Long, w As DAO.Workspace, started As Boolean
    On Error GoTo Failed
    Set f = Forms!frmUXComparison: id = Nz(f!lstTraits, 0)
    If id = 0 Then Err.Raise vbObjectError + 4700, , "Choose a member in the comparison grid."
    revision = Nz(f!lstTraits.Column(11), 0)
    Set w = DBEngine.Workspaces(0): w.BeginTrans: started = True
    For block = 33 To 39
        before = MajorRevisionValue(revision, "Trait" & block): value = f.Controls("cboTrait" & block).Value
        If Not SameFieldValue(before, value) Then ProposeMeetingField f!txtMeeting, id, "Trait" & block, value: count = count + 1
    Next
    before = MajorRevisionValue(revision, "PromotionRecommendation"): value = f!cboPromotion.Value
    If Not SameFieldValue(before, value) Then ProposeMeetingField f!txtMeeting, id, "PromotionRecommendation", value: count = count + 1
    w.CommitTrans: started = False
    UXLoadComparison
    f!lblStatus.Caption = count & " proposals staged. Saved reports are unchanged. Select one proposal per field below, then Apply."
    If UXHasControl(f, "uxPreviewPage") And count > 0 Then UXComparisonPage "Preview"
    UXStageComparison = True: Exit Function
Failed:
    If started Then w.Rollback
    Forms!frmUXComparison!lblStatus.Caption = Err.Description
End Function
Public Function UXApplyComparison() As Boolean
    Dim item As Variant, csv As String
    On Error GoTo Failed
    For Each item In Forms!frmUXComparison!lstProposals.ItemsSelected
        If Len(csv) > 0 Then csv = csv & ","
        csv = csv & Forms!frmUXComparison!lstProposals.ItemData(item)
    Next
    ApplyMeetingProposals csv
    Forms!frmUXComparison!lblStatus.Caption = "Selected proposals applied atomically. This frozen comparison is now historical; create a new snapshot for further edits."
    Forms!frmUXComparison!cmdStage.Enabled = False
    UXApplyComparison = True: Exit Function
Failed:
    Forms!frmUXComparison!lblStatus.Caption = Err.Description & " No partial proposal batch was applied."
End Function
Public Function UXEditorReview(ByVal mode As String) As Boolean
    Dim decision As String, f As Access.Form, i As Long, id As Long
    If EditorHasPending() Then
        decision = UXAsk("Review")
        If decision = "Save" Then
            If Not UXSaveEditor() Then Exit Function
        ElseIf decision <> "Discard" Then
            Exit Function
        End If
    End If
    id = Nz(Forms!frmEditor!txtReportID, 0)
    If id = 0 Then MsgBox "Save this new report before starting saved-revision review.": Exit Function
    DoCmd.OpenForm "frmMajorWorkbench"
    Set f = Forms!frmMajorWorkbench
    For i = 0 To f!lstReports.ListCount - 1: f!lstReports.Selected(i) = (Nz(f!lstReports.Column(0, i), 0) = id): Next
    Call UXWorkbenchMode(mode)
    UXEditorReview = True
End Function
Public Function UXStartGroupBulk() As Boolean
    Dim r As DAO.Recordset, key As Variant, values As Object, f As Access.Form, seen As Boolean, base As Object, db As DAO.Database
    groupContext = Nz(Forms!frmMajorWorkbench!cboGroup, 0)
    If groupContext = 0 Then MsgBox "Select a group first.": Exit Function
    bulkCSV = "": Set bulkChanges = CreateObject("Scripting.Dictionary"): Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & groupContext & " ORDER BY ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        If Len(bulkCSV) > 0 Then bulkCSV = bulkCSV & ","
        bulkCSV = bulkCSV & r!ReportID
        Set values = SavedFields(r!ReportID)
        If Not seen Then
            Set base = values: seen = True
        Else
            For Each key In base.Keys
                If Not SameFieldValue(base(key), values(key)) Then base(key) = Null
            Next
        End If
        r.MoveNext
    Loop
    r.Close
    If Not seen Then MsgBox "Add group members first.": Exit Function
    DoCmd.OpenForm "frmUXBulk"
    Set f = Forms!frmUXBulk
    For Each key In UXSharedKeys()
        If UXHasControl(f, "uxValue" & key) Then
            If f.Controls("uxValue" & key).ControlType = acCheckBox Then
                If IsNull(base(key)) Or Len(Nz(base(key), "")) = 0 Then
                    f.Controls("uxValue" & key) = Null
                ElseIf key = "ObservationStatus" Then
                    f.Controls("uxValue" & key) = (base(key) = "NOB")
                Else
                    f.Controls("uxValue" & key) = (base(key) = "Yes")
                End If
            Else
                f.Controls("uxValue" & key) = base(key)
            End If
        End If
    Next
    f!uxGroupTitle.Caption = Nz(MajorScalar("SELECT Title FROM SummaryGroup WHERE GroupID=" & groupContext), "Summary group")
    Call UXBulkSection(0)
    UXStartGroupBulk = True
End Function
Public Function UXBulkChanged(ByVal key As String) As Boolean
    If bulkChanges Is Nothing Then Set bulkChanges = CreateObject("Scripting.Dictionary")
    Dim c As Access.Control, value As Variant
    Set c = Forms!frmUXBulk.Controls("uxValue" & key): value = c.Value
    If c.ControlType = acCheckBox And Not IsNull(value) Then
        If key = "ObservationStatus" Then
            If value Then value = "NOB" Else value = "Observed"
        Else
            If value Then value = "Yes" Else value = "No"
        End If
    End If
    bulkChanges(key) = value
    Forms!frmUXBulk!uxStatus.Caption = bulkChanges.Count & " changed entries staged; no report revisions saved."
    UXBulkChanged = True
End Function
Public Function UXBulkSection(ByVal section As Long, Optional ByVal page As Long = 0) As Boolean
    Dim c As Access.Control, tags As Variant, f As Access.Form
    Set f = Forms!frmUXBulk: f!cmdPreview.SetFocus
    For Each c In f.Controls
        If Left$(c.Tag, 4) = "BLX:" Then tags = Split(c.Tag, ":"): c.Visible = (CLng(tags(1)) = section And CLng(tags(2)) = page)
    Next
    f!uxSection.Caption = Replace(UXSectionName(section), "&", "&&")
    UXBulkSection = True
End Function
Public Function UXPreviewBulk() As Boolean
    On Error GoTo Failed
    bulkBatch = PreviewReportChanges(bulkCSV, bulkChanges, Nz(Forms!frmUXBulk!txtReason, ""))
    Forms!frmUXBulk!lstPreview.RowSource = "SELECT i.ItemID,r.MemberName,UXFieldLabel(i.FieldKey),i.BeforeValue,i.NewValue FROM ChangeItem AS i INNER JOIN Report AS r ON i.ReportID=r.ReportID WHERE i.BatchID=" & bulkBatch & " ORDER BY r.MemberName,i.ItemID"
    Forms!frmUXBulk!lstPreview.Requery
    Forms!frmUXBulk!uxStatus.Caption = "Preview only. Review each complete before/after value before applying."
    UXPreviewBulk = True: Exit Function
Failed:
    Forms!frmUXBulk!uxStatus.Caption = Err.Description
End Function
Public Function UXApplyBulk() As Boolean
    On Error GoTo Failed
    ApplyFieldBatch bulkBatch
    Forms!frmUXBulk!uxStatus.Caption = "Applied reviewed block changes. One revision per report; previous values retained for guarded undo."
    UXApplyBulk = True: Exit Function
Failed:
    Forms!frmUXBulk!uxStatus.Caption = Err.Description & " No partial batch was applied."
End Function
Public Function UXBulkDetail() As Boolean
    Dim id As Long, r As DAO.Recordset, db As DAO.Database
    id = Nz(Forms!frmUXBulk!lstPreview, 0): If id = 0 Then Exit Function
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT * FROM ChangeItem WHERE ItemID=" & id, dbOpenSnapshot)
    Forms!frmUXBulk!txtFullValues = UXFieldLabel(r!FieldKey) & vbCrLf & "BEFORE:" & vbCrLf & Nz(r!BeforeValue, "(not entered)") & vbCrLf & "AFTER:" & vbCrLf & Nz(r!NewValue, "(clear value)")
    r.Close: UXBulkDetail = True
End Function
Public Function UXLoadContext() As Boolean
    Dim f As Access.Form, context As String
    context = Replace(Nz(DLookup("CommandName", "Workspace"), "Swift NAVFIT") & " | Contributor: " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & Nz(DLookup("ActiveProfileID", "Workspace"), 0)), ""), "&", "&&")
    For Each f In Forms
        If UXHasControl(f, "uxContext") Then f!uxContext.Caption = context
    Next
    UXLoadContext = True
End Function

Public Function UXCloseScreen(ByVal formName As String) As Boolean
    DoCmd.Close acForm, formName, acSaveNo
    UXCloseScreen = True
End Function

Public Function UXLoadWorkbench() As Boolean
    If Not LoadMajorWorkbench() Then Exit Function
    Call UXWorkbenchMode("Members")
    UXLoadWorkbench = True
End Function
