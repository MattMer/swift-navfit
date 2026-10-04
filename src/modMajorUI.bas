Attribute VB_Name = "modMajorUI"
Option Compare Binary
Option Explicit
Public Function LoadMajorWorkbench() As Boolean
    Dim f As Access.Form, key As Variant, keys As String
    Set f = Forms!frmMajorWorkbench
    f!lstReports.RowSource = "SELECT ReportID,Title,MemberName,FormKind FROM Report WHERE IsArchived=False ORDER BY MemberName,ReportID"
    f!cboGroup.RowSource = "SELECT GroupID,Title FROM SummaryGroup WHERE IsArchived=False ORDER BY Title"
    f!cboSnapshot.RowSource = "SELECT SnapshotID,'Group ' & GroupID & ' / ' & Format(AttestedAt,'yyyy-mm-dd hh:nn') FROM GroupSnapshot ORDER BY SnapshotID DESC"
    f!cboTemplate.RowSource = "SELECT TemplateID,Title FROM RouteTemplate WHERE IsArchived=False ORDER BY Title"
    f!cboRun.RowSource = "SELECT RunID,'Report ' & ReportID & ' / ' & RunState FROM RouteRun ORDER BY RunID DESC"
    f!cboMeeting.RowSource = "SELECT SetID,Title FROM ComparisonSet ORDER BY SetID DESC"
    f!cboBatch.RowSource = "SELECT BatchID,BatchKind & ' / ' & BatchState & ' / ' & Left(BatchNote,60) FROM ChangeBatch ORDER BY BatchID DESC"
    f!cboProposal.RowSource = "SELECT ProposalID,'Report ' & ReportID & ' / ' & FieldKey & ' / ' & ProposalState FROM MeetingProposal ORDER BY ProposalID DESC"
    f!cboFolder.RowSource = "SELECT FolderID,FolderName FROM Folder ORDER BY FolderName"
    f!cboReviewer.RowSource = "SELECT UserProfileID,DisplayName FROM UserProfile ORDER BY DisplayName,UserProfileID"
    For Each key In DraftFieldKeys()
        If key <> "FormKind" Then
            If Len(keys) > 0 Then keys = keys & ";"
            keys = keys & key & ";" & Replace(UXFieldLabel(CStr(key)), ";", ",")
        End If
    Next
    f!cboField.RowSource = keys
    f!cboField.ColumnCount = 2: f!cboField.ColumnWidths = "0;3000"
    LoadMajorWorkbench = True
End Function
Public Function ShowMajorDetails() As Boolean
    Dim f As Access.Form, r As DAO.Recordset, db As DAO.Database, i As Long, text As String, key As Variant, revision As Long, reportID As Long, row As Long, selected As Variant
    On Error GoTo Failed
    Set f = Forms!frmMajorWorkbench
    row = f!lstDetails.ListIndex
    If row < 0 Then
        For Each selected In f!lstDetails.ItemsSelected: row = CLng(selected): Exit For: Next
    End If
    If row < 0 Then Exit Function
    Set db = CurrentDb
    If f!lstDetails.RowSource Like "SELECT ReportID,FieldKey,*" Then
        Set r = db.OpenRecordset("SELECT * FROM ChangeItem WHERE BatchID=" & Nz(f!cboBatch, 0) & " AND ReportID=" & CLng(f!lstDetails.Column(0, row)) & " AND FieldKey=" & QuoteSQL(CStr(f!lstDetails.Column(1, row))), dbOpenSnapshot)
    ElseIf f!lstDetails.RowSource Like "SELECT ProposalID,*" Then
        Set r = db.OpenRecordset("SELECT * FROM MeetingProposal WHERE ProposalID=" & CLng(f!lstDetails.Column(0, row)), dbOpenSnapshot)
    ElseIf f!lstDetails.RowSource Like "SELECT StepOrdinal,*" Then
        Set r = db.OpenRecordset("SELECT * FROM RouteRunStep WHERE RunID=" & Nz(f!cboRun, 0) & " AND StepOrdinal=" & CLng(f!lstDetails.Column(0, row)), dbOpenSnapshot)
    ElseIf f!lstDetails.RowSource Like "SELECT ReportID,RevisionID,*" Then
        revision = CLng(f!lstDetails.Column(1, row))
        For Each key In DraftFieldKeys()
            If MajorRevisionHasField(revision, CStr(key)) Then text = text & key & ": " & Nz(MajorRevisionValue(revision, CStr(key)), "[null]") & vbCrLf Else text = text & key & ": [not recorded in this snapshot]" & vbCrLf
        Next
        f!txtDetail = text: ShowMajorDetails = True: Exit Function
    Else
        For i = 0 To 5: text = text & Nz(f!lstDetails.Column(i, row), "[null]") & vbCrLf: Next
        f!txtDetail = text: ShowMajorDetails = True: Exit Function
    End If
    If Not r.EOF Then
        For i = 0 To r.Fields.Count - 1: text = text & r.Fields(i).Name & ": " & Nz(r.Fields(i).Value, "[null]") & vbCrLf: Next
    End If
    r.Close: f!txtDetail = text: ShowMajorDetails = True: Exit Function
Failed:
    Forms!frmMajorWorkbench!lblStatus.Caption = Err.Description
End Function
Private Function WorkbenchReports() As String
    Dim item As Variant, f As Access.Form
    Set f = Forms!frmMajorWorkbench
    For Each item In f!lstReports.ItemsSelected
        If Len(WorkbenchReports) > 0 Then WorkbenchReports = WorkbenchReports & ","
        WorkbenchReports = WorkbenchReports & f!lstReports.ItemData(item)
    Next
    If Len(WorkbenchReports) = 0 Then Err.Raise vbObjectError + 4600, , "Select reports in the list."
End Function
Private Function WorkbenchOneReport() As Long
    Dim ids As Object, key As Variant
    Set ids = MajorIDs(WorkbenchReports())
    If ids.Count <> 1 Then Err.Raise vbObjectError + 4601, , "Select exactly one report."
    For Each key In ids.Keys: WorkbenchOneReport = ids(key): Next
End Function
Public Function MajorAction(ByVal action As String) As Boolean
    Dim f As Access.Form, id As Long, ids As Object, key As Variant, value As Variant, scopeID As Long, result As Object, w As DAO.Workspace, trans As Boolean, message As String
    On Error GoTo Failed
    Set f = Forms!frmMajorWorkbench
    If f!chkNull Then value = Null Else value = Nz(f!txtValue, "")
    Select Case action
        Case "CreateGroup"
            id = CreateSummaryGroup(Nz(f!txtTitle, ""), Nz(f!cboKind, ""), Nz(f!cboPaygrade, "")): f!cboGroup = id
        Case "AddMembers"
            Set ids = MajorIDs(WorkbenchReports())
            Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
            For Each key In ids.Keys: AddGroupMember Nz(f!cboGroup, 0), ids(key), Nz(f!cboPaygrade, ""), Nz(f!txtNote, ""): Next
            w.CommitTrans: trans = False
        Case "RemoveMember"
            RemoveGroupMember Nz(f!cboGroup, 0), WorkbenchOneReport()
        Case "FreezeGroup"
            id = FreezeSummaryGroup(Nz(f!cboGroup, 0), Nz(f!txtNote, "")): f!cboSnapshot = id
        Case "PreviewBatch"
            id = PreviewFieldBatch(WorkbenchReports(), Nz(f!cboField, ""), value, Nz(f!txtNote, "")): f!cboBatch = id
        Case "ShowBatch"
            f!lstDetails.RowSource = "SELECT ReportID,FieldKey,BeforeValue,NewValue,BeforeIsNull,NewIsNull FROM ChangeItem WHERE BatchID=" & Nz(f!cboBatch, 0) & " ORDER BY ReportID,ItemID"
        Case "ApplyBatch": ApplyFieldBatch Nz(f!cboBatch, 0)
        Case "PreviewUndo": id = PreviewBatchUndo(Nz(f!cboBatch, 0)): f!cboBatch = id
        Case "SaveDefault"
            Select Case Nz(f!cboScope, "")
                Case "Command": scopeID = 1
                Case "Folder": scopeID = Nz(f!cboFolder, 0)
                Case "Group": scopeID = Nz(f!cboGroup, 0)
                Case "Report": scopeID = WorkbenchOneReport()
                Case Else: Err.Raise vbObjectError + 4602, , "Select a default scope."
            End Select
            SetFieldDefault f!cboScope, scopeID, Nz(f!cboField, ""), value
        Case "SuggestDefault"
            Set result = ResolveFieldDefault(WorkbenchOneReport(), Nz(f!cboField, ""), Nz(f!cboGroup, 0))
            f!txtValue = result("Value"): f!chkNull = IsNull(result("Value"))
            f!lblStatus.Caption = "Suggestion from " & result("Source") & ". Preview a batch to apply it.": MajorAction = True: Exit Function
        Case "CreateRoute"
            id = CreateRouteTemplate(Nz(f!txtTitle, ""), Nz(f!txtReviewers, ""), Nz(f!txtNote, "")): f!cboTemplate = id
        Case "AddReviewer"
            id = Nz(f!cboReviewer, 0)
            If id = 0 Then Err.Raise vbObjectError + 4605, , "Choose a contributor for this step."
            If Len(Nz(f!txtReviewers, "")) > 0 Then f!txtReviewers = f!txtReviewers & ","
            f!txtReviewers = Nz(f!txtReviewers, "") & id
            If Len(Nz(f!txtReviewerNames, "")) > 0 Then f!txtReviewerNames = f!txtReviewerNames & " > "
            f!txtReviewerNames = Nz(f!txtReviewerNames, "") & Nz(MajorScalar("SELECT DisplayName FROM UserProfile WHERE UserProfileID=" & id), "")
        Case "ClearReviewers": f!txtReviewers = Null: f!txtReviewerNames = Null
        Case "StartRoute": id = StartRoute(WorkbenchOneReport(), Nz(f!cboTemplate, 0), Nz(f!cboGroup, 0)): f!cboRun = id
        Case "AcceptStep": DecideRouteStep Nz(f!cboRun, 0), "Accepted", Nz(f!txtNote, "")
        Case "ReturnStep": DecideRouteStep Nz(f!cboRun, 0), "Returned", Nz(f!txtNote, "")
        Case "CancelRoute": CancelRoute Nz(f!cboRun, 0)
        Case "ApproveRoute": ApproveSavedRoute Nz(f!cboRun, 0), Nz(f!txtNote, "")
        Case "ReopenApproval": ReopenApprovedReport WorkbenchOneReport(), Nz(f!txtNote, "")
        Case "ShowRoute"
            f!lstDetails.RowSource = "SELECT StepOrdinal,AssigneeID,StepState,Instructions,DecisionNote,DecisionAt FROM RouteRunStep WHERE RunID=" & Nz(f!cboRun, 0) & " ORDER BY StepOrdinal"
            f!lblStatus.Caption = "Route current: " & RouteIsCurrent(Nz(f!cboRun, 0)): MajorAction = True: Exit Function
        Case "CreateMeeting": id = CreateComparisonSet(Nz(f!cboSnapshot, 0), Nz(f!txtTitle, "")): f!cboMeeting = id
        Case "ShowMeeting"
            f!lstDetails.RowSource = "SELECT ReportID,RevisionID,SortOrdinal,MajorRevisionValue(RevisionID,'MemberName') AS Member,MajorRevisionValue(RevisionID,'Title') AS Title,MajorRevisionValue(RevisionID,'FormKind') AS Kind FROM ComparisonMember WHERE SetID=" & Nz(f!cboMeeting, 0) & " ORDER BY SortOrdinal"
        Case "ProposeField": id = ProposeMeetingField(Nz(f!cboMeeting, 0), WorkbenchOneReport(), Nz(f!cboField, ""), value): f!cboProposal = id
        Case "ApplyProposal": ApplyMeetingProposal Nz(f!cboProposal, 0)
        Case "OrderMeeting": SetMeetingOrder Nz(f!cboMeeting, 0), Nz(f!txtMeetingOrder, "")
        Case "ShowProposals"
            f!lstDetails.RowSource = "SELECT ProposalID,ReportID,FieldKey,ValueText,ProposalState,BaseRevision FROM MeetingProposal WHERE SetID=" & Nz(f!cboMeeting, 0) & " ORDER BY ProposalID"
        Case "ApplySelectedProposals"
            value = ""
            For Each key In f!lstDetails.ItemsSelected
                If Len(Nz(value, "")) > 0 Then value = value & ","
                value = value & f!lstDetails.ItemData(key)
            Next
            If f!lstDetails.RowSource Like "SELECT ProposalID,*" Then ApplyMeetingProposals CStr(value) Else Err.Raise vbObjectError + 4604, , "Show and select proposal rows first."
        Case "RunChecks"
            id = RunRulePack(WorkbenchOneReport(), InstallNavyDraftChecks())
            f!lstDetails.RowSource = "SELECT RuleID,FindingState,Message,ValidationRunID,'','' FROM ValidationFinding WHERE ValidationRunID=" & id & " ORDER BY FindingID"
            f!lblStatus.Caption = "Saved-revision draft checks " & id & ". Limited coverage; findings are not export approval.": MajorAction = True: Exit Function
        Case Else: Err.Raise vbObjectError + 4603, , "Unknown action."
    End Select
    LoadMajorWorkbench
    f!lblStatus.Caption = "Completed: " & action & IIf(id > 0, " / " & id, "")
    MajorAction = True: Exit Function
Failed:
    message = Err.Description
    If trans Then w.Rollback
    Forms!frmMajorWorkbench!lblStatus.Caption = message
End Function
