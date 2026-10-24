Attribute VB_Name = "modUXGroupOrganization"
Option Explicit
Private moveToken As String
Private moveOwner As Long
Public Function UXRefreshGroupFilter() As Boolean
    Dim f As Access.Form, sql As String
    Set f = Forms!frmLibrary
    sql = "SELECT 0 AS ID,'All reports' AS Label FROM AppMetadata UNION ALL SELECT -1,'Ungrouped' FROM AppMetadata UNION ALL SELECT GroupID,Title FROM SummaryGroup ORDER BY Label"
    f!lstFolders.RowSource = sql: f!lstFolders.Requery
    f!uxFolderFilter.RowSource = sql: f!uxFolderFilter.Requery
    If IsNull(f!lstFolders) Then f!lstFolders = 0
    f!uxFolderFilter = f!lstFolders
    UXRefreshGroupFilter = True
End Function
Private Function MoveStamp(ByVal id As Long, ByVal destination As Long) As String
    Dim r As DAO.Recordset, db As DAO.Database
    MoveStamp = CStr(CurrentRevision(id)) & ":" & Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") & ":" & CStr(Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True))
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT GroupID FROM GroupMember WHERE ReportID=" & id & " ORDER BY GroupID", dbOpenSnapshot)
    Do While Not r.EOF: MoveStamp = MoveStamp & "|" & r!GroupID & ":" & GroupStamp(r!GroupID): r.MoveNext: Loop
    r.Close
    If destination > 0 Then MoveStamp = MoveStamp & "|destination:" & destination & ":" & GroupStamp(destination)
End Function
Private Sub CheckMove(ByVal id As Long, ByVal destination As Long)
    Dim grade As String
    If Nz(DLookup("IsArchived", "Report", "ReportID=" & id), True) Then Err.Raise vbObjectError + 6800, "modUXGroupOrganization.CheckMove", "Recover the report before moving it."
    If Nz(DLookup("WorkflowState", "Report", "ReportID=" & id), "") <> "Draft" Then Err.Raise vbObjectError + 6801, "modUXGroupOrganization.CheckMove", "Reopen local approval before moving this report."
    If destination = 0 Then Exit Sub
    If Nz(DLookup("IsArchived", "SummaryGroup", "GroupID=" & destination), True) Then Err.Raise vbObjectError + 6802, "modUXGroupOrganization.CheckMove", "Choose an active group."
    grade = ValidationWornPaygrade(Nz(DLookup("RankRate", "Report", "ReportID=" & id), ""))
    If grade = "" Or grade <> Nz(DLookup("Paygrade", "SummaryGroup", "GroupID=" & destination), "") Or Nz(DLookup("FormKind", "Report", "ReportID=" & id), "") <> Nz(DLookup("FormKind", "SummaryGroup", "GroupID=" & destination), "") Then Err.Raise vbObjectError + 6803, "modUXGroupOrganization.CheckMove", "Set the report's worn rank/rate and choose a matching report type/paygrade group."
End Sub
Public Sub MoveReportGroup(ByVal id As Long, ByVal destination As Long, ByVal expected As String)
    Dim w As DAO.Workspace, db As DAO.Database, r As DAO.Recordset, old As String, failure As Long, message As String, op As Long
    Set w = DBEngine.Workspaces(0): Set db = CurrentDb
    On Error GoTo Failed
    w.BeginTrans
    CheckMove id, destination
    If MoveStamp(id, destination) <> expected Then Err.Raise vbObjectError + 6804, "modUXGroupOrganization.MoveReportGroup", "Report or group changed. Preview the move again."
    old = UXTableGroups(id)
    Set r = db.OpenRecordset("SELECT GroupID FROM GroupMember WHERE ReportID=" & id & " ORDER BY GroupID", dbOpenSnapshot)
    Do While Not r.EOF
        If r!GroupID <> destination Then RemoveGroupMember r!GroupID, id
        r.MoveNext
    Loop
    r.Close
    If destination > 0 Then
        If DCount("*", "GroupMember", "ReportID=" & id & " AND GroupID=" & destination) = 0 Then AddGroupMember destination, id, DLookup("Paygrade", "SummaryGroup", "GroupID=" & destination)
    End If
    op = BeginOperation("Move report to summary group", "Report " & id & "; BEFORE: " & old & "; AFTER: " & UXTableGroups(id))
    FinishOperation op, "Completed", "Membership changed atomically. Saved report values unchanged; affected roster reviews require rechecking."
    w.CommitTrans: Exit Sub
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modUXGroupOrganization", "MoveReportGroup", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary: If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0: SwiftRaise swiftPrimary
End Sub
Public Function UXChooseReportGroup() As Boolean
    Dim id As Long
    id = Nz(Forms!frmLibrary!lstReports, 0): If id = 0 Then Exit Function
    If Not UXLeaveEditor() Then Exit Function
    moveToken = ""
    DoCmd.OpenForm "frmUXMoveGroup", , , , , , CStr(id)
    UXChooseReportGroup = True
End Function
Public Function UXLoadMoveGroup() As Boolean
    Call UXAssertView("frmUXMoveGroup")
    moveToken = "": moveOwner = 0
    With Forms!frmUXMoveGroup
        !cmdApply.Enabled = False: !txtPreview = Null
        !reportID = Val(Nz(.OpenArgs, "0"))
        !cboGroup.RowSource = "SELECT 0 AS ID,'Ungrouped' AS Label FROM AppMetadata UNION ALL SELECT GroupID,Title FROM SummaryGroup WHERE IsArchived=False ORDER BY Label"
        !cboGroup = 0
        !lblMember.Caption = UXMemberLabel(!reportID)
    End With
    UXLoadMoveGroup = True
End Function
Public Function UXPreviewMoveGroup() As Boolean
    Dim id As Long, destination As Long, csv As String, db As DAO.Database, r As DAO.Recordset
    Call UXAssertView("frmUXMoveGroup")
    moveToken = "": moveOwner = 0: Forms!frmUXMoveGroup!cmdApply.Enabled = False
    On Error GoTo Failed
    id = Forms!frmUXMoveGroup!reportID: destination = Nz(Forms!frmUXMoveGroup!cboGroup, 0)
    CheckMove id, destination
    Forms!frmUXMoveGroup!txtPreview = "BEFORE: " & UXTableGroups(id) & vbCrLf & "AFTER: " & IIf(destination = 0, "Ungrouped", Nz(DLookup("Title", "SummaryGroup", "GroupID=" & destination), "")) & vbCrLf & "The report leaves its other groups. Report field values remain unchanged. Affected roster reviews become stale."
    If destination > 0 Then
        csv = CStr(id): Set db = CurrentDb
        Set r = db.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & destination & " AND ReportID<>" & id, dbOpenSnapshot)
        Do While Not r.EOF: csv = csv & "," & r!ReportID: r.MoveNext: Loop
        r.Close
        Forms!frmUXMoveGroup!txtPreview = Forms!frmUXMoveGroup!txtPreview & vbCrLf & GroupImpactText(destination, , csv)
    End If
    moveOwner = ActiveContributorID(): moveToken = MoveStamp(id, destination): Forms!frmUXMoveGroup!cmdApply.Enabled = True: UXPreviewMoveGroup = True: Exit Function
Failed:
    moveToken = "": Forms!frmUXMoveGroup!lblStatus.Caption = Err.Description
End Function
Public Function UXApplyMoveGroup() As Boolean
    Dim committed As Boolean, message As String
    Call UXAssertView("frmUXMoveGroup")
    On Error GoTo Failed
    If Len(moveToken) = 0 Then Err.Raise vbObjectError + 6805, "modUXGroupOrganization.UXApplyMoveGroup", "Preview this move first."
    If moveOwner <> ActiveContributorID() Then Err.Raise vbObjectError + 6805, "modUXGroupOrganization.UXApplyMoveGroup", "Preview this move as the current contributor before Apply."
    MoveReportGroup CLng(Forms!frmUXMoveGroup!reportID), CLng(Nz(Forms!frmUXMoveGroup!cboGroup, 0)), moveToken
    committed = True
    moveToken = "": Forms!frmUXMoveGroup!lblStatus.Caption = "Moved successfully. Saved report values unchanged."
    Forms!frmUXMoveGroup!cmdApply.Enabled = False
    RefreshLibrary
    UXApplyMoveGroup = True: Exit Function
Failed:
    message = Err.Description: moveToken = "": moveOwner = 0
    Forms!frmUXMoveGroup!cmdApply.Enabled = False
    If committed Then message = "Move completed; the display update failed. Reopen Reports to inspect the saved membership. " & message
    Forms!frmUXMoveGroup!lblStatus.Caption = message
End Function
Public Function UXSetGroupWorkflow() As Boolean
    On Error GoTo Failed
    With Forms!frmUXGroups
        SetGroupWorkflow CLng(Nz(!lstGroups, 0)), Nz(!uxStateStamp, ""), Nz(!uxGroupState, "")
    End With
    UXGroupList
    UXSetGroupWorkflow = True: Exit Function
Failed:
    Forms!frmUXGroups!uxLifecycleStatus.Caption = Err.Description
End Function
