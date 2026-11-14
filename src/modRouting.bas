Attribute VB_Name = "modRouting"
Option Compare Database
Option Explicit
Private Function RouteScalar(ByVal sql As String) As Variant
    Dim db As DAO.Database
    Set db = CurrentDb
    RouteScalar = RouteValue(db, sql)
End Function
Private Function RouteValue(ByVal db As DAO.Database, ByVal sql As String) As Variant
    Dim r As DAO.Recordset
    Set r = db.OpenRecordset(sql, dbOpenSnapshot)
    If r.EOF Then RouteValue = Null Else RouteValue = r.Fields(0).Value
    r.Close: Set r = Nothing
End Function
Private Function RouteProfile() As Long
    RouteProfile = ActiveContributorID()
End Function
Public Function AddContributor(ByVal name As String, Optional ByVal memberID As Long = 0) As Long
    Dim r As DAO.Recordset, retired As Variant
    Call RouteProfile
    name = Trim$(name)
    If Len(name) = 0 Or Len(name) > 120 Then Err.Raise vbObjectError + 1801, "modRouting.AddContributor", "Enter a contributor name of 1 to 120 characters."
    If DCount("*", "UserProfile", "IsRetired=False AND DisplayName=" & QuoteSQL(name)) > 0 Then Err.Raise vbObjectError + 1802, "modRouting.AddContributor", "A contributor with this display name already exists."
    'A deleted contributor added again gets their old profile back, with its history.
    retired = MajorScalar("SELECT Min(UserProfileID) FROM UserProfile WHERE IsRetired=True AND DisplayName=" & QuoteSQL(name))
    If memberID > 0 And IsNull(retired) Then retired = MajorScalar("SELECT Min(UserProfileID) FROM UserProfile WHERE IsRetired=True AND MemberID=" & memberID)
    If IsNull(retired) Then
        Set r = CurrentDb.OpenRecordset("UserProfile", dbOpenDynaset)
        r.AddNew: r!ProfileGuid = NewIdentity(): r!IsRetired = False
    Else
        Set r = CurrentDb.OpenRecordset("SELECT * FROM UserProfile WHERE UserProfileID=" & retired, dbOpenDynaset)
        r.Edit: r!IsRetired = False
    End If
    r!DisplayName = name
    If memberID > 0 Then r!MemberID = memberID
    r.Update: r.Bookmark = r.LastModified
    AddContributor = r!UserProfileID: r.Close: Set r = Nothing
End Function
'A contributor from the command directory: "RANK LAST, FIRST MI", linked to the member.
Public Function AddContributorFromMember(ByVal memberID As Long) As Long
    Dim r As DAO.Recordset, name As String
    Set r = CurrentDb.OpenRecordset("SELECT * FROM CommandMember WHERE MemberID=" & memberID, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 1832, "modRouting.AddContributorFromMember", "Choose a member from the directory."
    If r!IsArchived Then r.Close: Err.Raise vbObjectError + 1832, "modRouting.AddContributorFromMember", "Recover the member before adding them as a contributor."
    name = Nz(r!ReportName, "")
    If Len(name) = 0 Then name = NavyNameFormat(r!FullName)
    name = Trim$(Nz(r!RankRate, "") & " " & name): r.Close
    If DCount("*", "UserProfile", "IsRetired=False AND MemberID=" & memberID) > 0 Then Err.Raise vbObjectError + 1833, "modRouting.AddContributorFromMember", name & " is already a contributor."
    AddContributorFromMember = AddContributor(name, memberID)
End Function
'Deleting a contributor keeps their name on everything they did, hides them from
'every contributor list, takes them out of route templates and active review
'routes, and cancels the review requests assigned to them. Returns a summary.
Public Function DeleteContributor(ByVal profileID As Long) As String
    Dim db As DAO.Database, w As DAO.Workspace, r As DAO.Recordset, name As String, tasks As Long, routes As String, trans As Boolean, operation As Long
    On Error GoTo Failed
    If profileID = RouteProfile() Then Err.Raise vbObjectError + 1830, "modRouting.DeleteContributor", "Switch to another contributor before deleting this one."
    If Nz(MajorScalar("SELECT IsRetired FROM UserProfile WHERE UserProfileID=" & profileID), True) Then Err.Raise vbObjectError + 1831, "modRouting.DeleteContributor", "Choose a contributor to delete."
    name = MajorScalar("SELECT DisplayName FROM UserProfile WHERE UserProfileID=" & profileID)
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    Set r = db.OpenRecordset("SELECT * FROM ReviewTask WHERE AssignedTo=" & profileID & " AND TaskState='Pending'", dbOpenDynaset)
    Do While Not r.EOF
        r.Edit: r!TaskState = "Cancelled": r!FinishedAt = Now(): r.Update
        Call RouteEvent(db, r!TaskID, RouteProfile(), "Cancelled", "Review request removed: " & name & " was deleted as a contributor.")
        tasks = tasks + 1: r.MoveNext
    Loop
    r.Close: Set r = Nothing
    routes = RouteRemoveContributor(db, profileID, name)
    db.Execute "UPDATE UserProfile SET IsRetired=True WHERE UserProfileID=" & profileID, dbFailOnError
    DeleteContributor = "Deleted " & name & ". " & routes & " " & tasks & " review request(s) cancelled. Their past work keeps their name."
    operation = BeginOperation("Delete contributor", name)
    FinishOperation operation, "Completed", DeleteContributor
    w.CommitTrans: trans = False: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRouting", "DeleteContributor", "Execute", Err.Number, Err.Description, Err.Source, IIf(trans, "Active", "Not active"))
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r": Set r = Nothing
    If trans Then SwiftRollback w, swiftPrimary
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
Public Sub RequireContributorSwitch()
    If UXGridPendingCount() > 0 Then Err.Raise vbObjectError + 4815, "modRouting.RequireContributorSwitch", "Retry or discard pending comparison cells before switching contributors."
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 1803, "modRouting.RequireContributorSwitch", "Close the draft editor and review comments before switching contributors."
    If DCount("*", "Workspace", "WorkspaceID=1") <> 1 Then Err.Raise 5, "modRouting.RequireContributorSwitch", "Complete workspace setup first."
End Sub
Public Function SwitchContributor(ByVal profileID As Long) As Boolean
    RequireContributorSwitch
    If DCount("*", "UserProfile", "IsRetired=False AND UserProfileID=" & profileID) = 0 Then Err.Raise vbObjectError + 1804, "modRouting.SwitchContributor", "Select an existing contributor."
    'The shared row only remembers the last choice; another session saving at the same
    'moment must not stop this one.
    On Error Resume Next
    CurrentDb.Execute "UPDATE Workspace SET ActiveProfileID=" & profileID & " WHERE WorkspaceID=1", dbFailOnError
    On Error GoTo 0
    BindActiveContributor profileID
    If CurrentProject.AllForms("frmLibrary").IsLoaded Then Call UXReportFilterChoices
    UXLoadContext
    SwitchContributor = True
End Function
Public Function AssignReview(ByVal reportID As Long, ByVal expectedRevision As Long, ByVal assigneeID As Long, ByVal due As Variant, ByVal instructions As String) As Long
    Dim r As DAO.Recordset, w As DAO.Workspace, db As DAO.Database, trans As Boolean, id As Long, pid As Long, failure As Long, message As String, stage As String
    On Error GoTo Failed
    pid = RouteProfile()
    If DCount("*", "UserProfile", "IsRetired=False AND UserProfileID=" & assigneeID) = 0 Then Err.Raise vbObjectError + 1804, "modRouting.AssignReview", "Select an existing reviewer."
    If Not IsNull(due) Then due = DateValue(due)
    stage = "Begin assignment"
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    Set r = db.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 1805, "modRouting.AssignReview", "Select a saved report."
    If r!IsArchived Or r!WorkflowState <> "Draft" Then Err.Raise vbObjectError + 1806, "modRouting.AssignReview", "Review assignment requires an active draft."
    If Nz(r!CurrentRevisionID, 0) <> expectedRevision Or expectedRevision = 0 Then Err.Raise vbObjectError + 1807, "modRouting.AssignReview", "Saved revision changed. Refresh before assigning."
    r.Close: Set r = Nothing
    RetainReportCheckpoint reportID, "Review assigned"
    If RouteValue(db, "SELECT Count(*) FROM ReviewTask WHERE ReportID=" & reportID & " AND TaskState='Pending'") > 0 Then Err.Raise vbObjectError + 1808, "modRouting.AssignReview", "Finish or cancel the pending review before assigning another."
    'Someone already in the report's active review route acts through that route; a
    'separate assignment would give them nothing they can act on.
    Dim inRoute As Variant
    inRoute = RouteValue(db, "SELECT Min(s.StepOrdinal) FROM RouteRunStep AS s INNER JOIN RouteRun AS u ON s.RunID=u.RunID WHERE u.ReportID=" & reportID & " AND u.RunState In ('Pending','Accepted') AND s.AssigneeID=" & assigneeID)
    If Not IsNull(inRoute) Then Err.Raise vbObjectError + 1817, "modRouting.AssignReview", Nz(RouteValue(db, "SELECT DisplayName FROM UserProfile WHERE UserProfileID=" & assigneeID), "This contributor") & " is already in this report's review route (step " & inRoute & "). They review it through the route, so no separate assignment is needed."
    stage = "Write task"
    Set r = db.OpenRecordset("ReviewTask", dbOpenDynaset)
    r.AddNew: r!TaskGuid = NewIdentity(): r!ReportID = reportID: r!RevisionID = expectedRevision
    r!RequestedBy = pid: r!AssignedTo = assigneeID: r!RequestedAt = Now(): r!DueDate = due
    r!Instructions = instructions: r!TaskState = "Pending": r.Update: r.Bookmark = r.LastModified: id = r!TaskID: r.Close: Set r = Nothing
    stage = "Write assignment event"
    Call RouteEvent(db, id, pid, "Assigned", instructions)
    stage = "Commit assignment"
    w.CommitTrans: trans = False: AssignReview = id: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRouting", "AssignReview", stage, Err.Number, Err.Description, Err.Source, IIf(trans, "Active", "Not active"))
    failure = CLng(swiftPrimary("Number")): message = stage & ": " & CStr(swiftPrimary("Description"))
    If trans Then SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r": Set r = Nothing
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
Private Sub RouteEvent(ByVal db As DAO.Database, ByVal taskID As Long, ByVal authorID As Long, ByVal kind As String, ByVal notes As String)
    Dim r As DAO.Recordset
    Set r = db.OpenRecordset("ReviewTaskEvent", dbOpenDynaset)
    r.AddNew: r!TaskID = taskID: r!AuthorProfileID = authorID: r!EventAt = Now(): r!EventKind = kind: r!Notes = notes: r.Update: r.Bookmark = r.LastModified
    AuditRecord "Review event", CLng(r!TaskEventID): r.Close: Set r = Nothing
End Sub
Public Function FinishReview(ByVal taskID As Long, ByVal outcome As String, ByVal notes As String) As Boolean
    Dim r As DAO.Recordset, w As DAO.Workspace, db As DAO.Database, pid As Long, trans As Boolean, failure As Long, message As String
    On Error GoTo Failed
    pid = RouteProfile()
    If outcome <> "Accepted" And outcome <> "Returned" And outcome <> "Cancelled" Then Err.Raise vbObjectError + 1809, "modRouting.FinishReview", "Choose an available review outcome."
    If outcome = "Returned" And Len(Trim$(Replace(Replace(Replace(notes, vbCr, ""), vbLf, ""), vbTab, ""))) = 0 Then Err.Raise vbObjectError + 1810, "modRouting.FinishReview", "Explain why the review is returned."
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    Set r = db.OpenRecordset("SELECT * FROM ReviewTask WHERE TaskID=" & taskID, dbOpenDynaset)
    If r.EOF Then Err.Raise vbObjectError + 1811, "modRouting.FinishReview", "Select a review assignment."
    If r!TaskState <> "Pending" Then Err.Raise vbObjectError + 1812, "modRouting.FinishReview", "This review already has an outcome."
    If outcome = "Cancelled" Then
        If pid <> r!RequestedBy Then Err.Raise vbObjectError + 1813, "modRouting.FinishReview", "Switch to the requesting contributor to cancel this assignment."
    Else
        If pid <> r!AssignedTo Then Err.Raise vbObjectError + 1814, "modRouting.FinishReview", "Switch to the assigned reviewer to record an outcome."
        If Nz(RouteValue(db, "SELECT IsArchived FROM Report WHERE ReportID=" & r!ReportID), True) Then Err.Raise vbObjectError + 1815, "modRouting.FinishReview", "Archived reports cannot receive review outcomes."
        If Nz(RouteValue(db, "SELECT WorkflowState FROM Report WHERE ReportID=" & r!ReportID), "") <> "Draft" Then Err.Raise vbObjectError + 1815, "modRouting.FinishReview", "Reopen local approval before recording review outcomes."
        'F01: the assignee holds the shared draft and may save; finishing records the revision.
    End If
    RetainReportCheckpoint CLng(r!ReportID), "Review " & outcome
    r.Edit: r!TaskState = outcome: r!FinishedAt = Now(): r!FinishedRevisionID = RouteValue(db, "SELECT CurrentRevisionID FROM Report WHERE ReportID=" & r!ReportID): r.Update: r.Close: Set r = Nothing
    Call RouteEvent(db, taskID, pid, outcome, notes)
    w.CommitTrans: trans = False: FinishReview = True: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRouting", "FinishReview", "Execute", Err.Number, Err.Description, Err.Source, IIf(trans, "Active", "Not active"))
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r": Set r = Nothing
    If trans Then SwiftRollback w, swiftPrimary
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
Public Function OpenMyWork(Optional ByVal reportID As Long = 0) As Boolean
    Dim wasLoaded As Boolean
    wasLoaded = CurrentProject.AllForms("frmMyWork").IsLoaded
    DoCmd.OpenForm "frmMyWork"
    Forms!frmMyWork!txtReportID = reportID
    Forms!frmMyWork!txtExpectedRevision = Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & reportID), 0)
    If reportID > 0 Then
        Forms!frmMyWork!lblAssignment.Caption = "Assign review: " & UXWorkReport(reportID) & ". Saved revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & Forms!frmMyWork!txtExpectedRevision), 0) & "."
    Else
        Forms!frmMyWork!lblAssignment.Caption = "Open a saved report to assign a review."
    End If
    Call LoadMyWork(wasLoaded): OpenMyWork = True
End Function
Public Function RoutingFromEditor() As Boolean
    RoutingFromEditor = OpenMyWork(Nz(Forms!frmEditor!txtReportID, 0))
End Function
Public Function LoadMyWork(Optional ByVal refreshQueue As Boolean = True) As Boolean
    Dim sql As String, pid As Long, reportID As Long
    pid = RouteProfile()
    reportID = Nz(Forms!frmMyWork!txtReportID, 0)
    Forms!frmMyWork!txtExpectedRevision = Nz(RouteScalar("SELECT CurrentRevisionID FROM Report WHERE ReportID=" & reportID), 0)
    If reportID > 0 Then
        Forms!frmMyWork!lblAssignment.Caption = "Assign review: " & UXWorkReport(reportID) & ". Saved revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & Forms!frmMyWork!txtExpectedRevision), 0) & "."
    Else
        Forms!frmMyWork!lblAssignment.Caption = "Open a saved report to assign a review."
    End If
    Forms!frmMyWork!lblProfile.Caption = "Active contributor: " & DLookup("DisplayName", "UserProfile", "UserProfileID=" & pid) & ". Profiles are attribution, not authentication."
    Forms!frmMyWork!cboReviewer.RowSource = "SELECT UserProfileID,DisplayName FROM UserProfile WHERE IsRetired=False ORDER BY DisplayName": Forms!frmMyWork!cboReviewer.Requery
    sql = "SELECT t.TaskID,UXIdentity(r.OfficialName,r.MemberName) & ' | ' & r.FormKind & ' | ' & Nz(Format(r.PeriodEnd,'yyyy-mm-dd'),'No period') AS [Member and report],p.DisplayName AS Reviewer,t.TaskState AS [Review step],t.DueDate AS [Due date],IIf(t.RevisionID=r.CurrentRevisionID,'Current','Stale') AS [Saved version],r.IsArchived FROM (ReviewTask AS t INNER JOIN Report AS r ON t.ReportID=r.ReportID) INNER JOIN UserProfile AS p ON t.AssignedTo=p.UserProfileID"
    If Not Nz(Forms!frmMyWork!chkAll, False) Then sql = sql & " WHERE t.AssignedTo=" & pid & " AND t.TaskState='Pending'"
    Forms!frmMyWork!lstTasks.RowSource = sql & " ORDER BY t.TaskID DESC": Forms!frmMyWork!lstTasks.Requery
    Forms!frmMyWork!txtHistory = Null
    If refreshQueue Then Call UXLoadWorkQueue
    LoadMyWork = True
End Function
Public Function TaskDetails() As Boolean
    Dim r As DAO.Recordset, id As Long, text As String
    id = Nz(Forms!frmMyWork!lstTasks, 0)
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ReviewTask WHERE TaskID=" & id, dbOpenSnapshot)
    If r.EOF Then r.Close: Set r = Nothing: Exit Function
    text = "Report " & r!ReportID & " / saved revision ID " & r!RevisionID & " / " & r!TaskState & vbCrLf & "Requested by " & DLookup("DisplayName", "UserProfile", "UserProfileID=" & r!RequestedBy) & " / " & r!RequestedAt & vbCrLf & "Due: " & Nz(r!DueDate, "No due date")
    If Not IsNull(r!DueDate) And r!TaskState = "Pending" Then
        If r!DueDate < Date Then text = text & " (overdue)"
    End If
    If Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & r!ReportID), 0) <> r!RevisionID Then text = text & vbCrLf & "STALE: saved report changed since assignment."
    r.Close: Set r = Nothing
    Set r = CurrentDb.OpenRecordset("SELECT e.*,p.DisplayName FROM ReviewTaskEvent AS e INNER JOIN UserProfile AS p ON e.AuthorProfileID=p.UserProfileID WHERE e.TaskID=" & id & " ORDER BY e.TaskEventID", dbOpenSnapshot)
    Do While Not r.EOF
        text = text & vbCrLf & vbCrLf & r!EventKind & " / " & r!DisplayName & " / " & r!EventAt & vbCrLf & Nz(r!Notes, "")
        r.MoveNext
    Loop
    r.Close: Set r = Nothing: Forms!frmMyWork!txtHistory = text: TaskDetails = True
End Function
'Resources > Contributors: the people who use this workspace.
Public Function UXContributorsLoad() As Boolean
    If Not CurrentProject.AllForms("frmUXContributors").IsLoaded Then Exit Function
    Forms!frmUXContributors!lstContributors.RowSource = "SELECT UserProfileID,DisplayName FROM UserProfile WHERE IsRetired=False ORDER BY DisplayName"
    Forms!frmUXContributors!lstContributors.Requery
    UXContributorsLoad = True
End Function
Public Function ContributorUI(ByVal action As String) As Boolean
    Dim f As Access.Form
    On Error GoTo Failed
    Set f = Forms!frmUXContributors
    If action = "Add" Then
        Call AddContributor(Nz(f!txtContributor, ""))
        f!txtContributor = Null
    ElseIf action = "Directory" Then
        DoCmd.OpenForm "frmUXSeniorPicker", , , , , , "contributor"
        ContributorUI = True: Exit Function
    ElseIf action = "Delete" Then
        If IsNull(f!lstContributors) Then Err.Raise vbObjectError + 1820, "modRouting.ContributorUI", "Select a contributor in the list first."
        If MsgBox("Delete " & f!lstContributors.Column(1) & "?" & vbCrLf & vbCrLf & "They are removed from route templates and active review routes, and review requests assigned to them are cancelled. Their past work keeps their name.", vbYesNo + vbExclamation, "Delete contributor") <> vbYes Then Exit Function
        MsgBox DeleteContributor(CLng(f!lstContributors)), vbInformation, "Delete contributor"
    Else
        If IsNull(f!lstContributors) Then Err.Raise vbObjectError + 1820, "modRouting.ContributorUI", "Select a contributor in the list first."
        Call SwitchContributor(Nz(f!lstContributors, 0))
    End If
    Call UXContributorsLoad
    If CurrentProject.AllForms("frmMyWork").IsLoaded Then Call LoadMyWork
    ContributorUI = True: Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Local contributors"
End Function
Public Function AssignReviewUI() As Boolean
    On Error GoTo Failed
    Call AssignReview(Nz(Forms!frmMyWork!txtReportID, 0), Nz(Forms!frmMyWork!txtExpectedRevision, 0), Nz(Forms!frmMyWork!cboReviewer, 0), Forms!frmMyWork!txtDue, Nz(Forms!frmMyWork!txtNotes, ""))
    Forms!frmMyWork!txtNotes = Null: Call LoadMyWork(): AssignReviewUI = True: Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Assign review"
End Function
Public Function FinishReviewUI(ByVal outcome As String) As Boolean
    On Error GoTo Failed
    FinishReviewUI = FinishReview(Nz(Forms!frmMyWork!lstTasks, 0), outcome, Nz(Forms!frmMyWork!txtNotes, ""))
    Forms!frmMyWork!txtNotes = Null: Call LoadMyWork(): Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Review outcome"
End Function
Public Function OpenReviewTaskDraft() As Boolean
 Dim id As Long, task As Long
 task = Nz(Forms!frmMyWork!lstTasks, 0)
 id = Nz(DLookup("ReportID", "ReviewTask", "TaskID=" & task), 0)
 If id > 0 Then OpenReviewTaskDraft = UXOpenReportTool(id, "Review", 0, 0, task)
End Function
