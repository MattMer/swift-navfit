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
    RouteProfile = Nz(RouteScalar("SELECT ActiveProfileID FROM Workspace WHERE WorkspaceID=1"), 0)
    If RouteProfile = 0 Then Err.Raise vbObjectError + 1800, , "Configure a contributor first."
End Function
Public Function AddContributor(ByVal name As String) As Long
    Dim r As DAO.Recordset
    Call RouteProfile
    name = Trim$(name)
    If Len(name) = 0 Or Len(name) > 120 Then Err.Raise vbObjectError + 1801, , "Enter a contributor name of 1 to 120 characters."
    If DCount("*", "UserProfile", "DisplayName='" & Replace(name, "'", "''") & "'") > 0 Then Err.Raise vbObjectError + 1802, , "A contributor with this display name already exists."
    Set r = CurrentDb.OpenRecordset("UserProfile", dbOpenDynaset)
    r.AddNew: r!ProfileGuid = NewIdentity(): r!DisplayName = name: r.Update: r.Bookmark = r.LastModified
    AddContributor = r!UserProfileID: r.Close: Set r = Nothing
End Function
Public Function SwitchContributor(ByVal profileID As Long) As Boolean
    If CurrentProject.AllForms("frmEditor").IsLoaded Or CurrentProject.AllForms("frmReviewComments").IsLoaded Then Err.Raise vbObjectError + 1803, , "Close the draft editor and review comments before switching contributors."
    If DCount("*", "UserProfile", "UserProfileID=" & profileID) = 0 Then Err.Raise vbObjectError + 1804, , "Select an existing contributor."
    CurrentDb.Execute "UPDATE Workspace SET ActiveProfileID=" & profileID & " WHERE WorkspaceID=1", dbFailOnError
    SwitchContributor = True
End Function
Public Function AssignReview(ByVal reportID As Long, ByVal expectedRevision As Long, ByVal assigneeID As Long, ByVal due As Variant, ByVal instructions As String) As Long
    Dim r As DAO.Recordset, w As DAO.Workspace, db As DAO.Database, trans As Boolean, id As Long, pid As Long, failure As Long, message As String, stage As String
    On Error GoTo Failed
    pid = RouteProfile()
    If DCount("*", "UserProfile", "UserProfileID=" & assigneeID) = 0 Then Err.Raise vbObjectError + 1804, , "Select an existing reviewer."
    If Not IsNull(due) Then due = DateValue(due)
    stage = "Begin assignment"
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    Set r = db.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 1805, , "Select a saved report."
    If r!IsArchived Or r!WorkflowState <> "Draft" Then Err.Raise vbObjectError + 1806, , "Review assignment requires an active draft."
    If Nz(r!CurrentRevisionID, 0) <> expectedRevision Or expectedRevision = 0 Then Err.Raise vbObjectError + 1807, , "Saved revision changed. Refresh before assigning."
    r.Close: Set r = Nothing
    If RouteValue(db, "SELECT Count(*) FROM ReviewTask WHERE ReportID=" & reportID & " AND TaskState='Pending'") > 0 Then Err.Raise vbObjectError + 1808, , "Finish or cancel the pending review before assigning another."
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
    failure = Err.Number: message = stage & ": " & Err.Description
    On Error Resume Next
    If trans Then w.Rollback
    If Not r Is Nothing Then r.Close: Set r = Nothing
    On Error GoTo 0
    Err.Raise failure, "AssignReview", message
End Function
Private Sub RouteEvent(ByVal db As DAO.Database, ByVal taskID As Long, ByVal authorID As Long, ByVal kind As String, ByVal notes As String)
    Dim r As DAO.Recordset
    Set r = db.OpenRecordset("ReviewTaskEvent", dbOpenDynaset)
    r.AddNew: r!TaskID = taskID: r!AuthorProfileID = authorID: r!EventAt = Now(): r!EventKind = kind: r!Notes = notes: r.Update: r.Close: Set r = Nothing
End Sub
Public Function FinishReview(ByVal taskID As Long, ByVal outcome As String, ByVal notes As String) As Boolean
    Dim r As DAO.Recordset, w As DAO.Workspace, db As DAO.Database, pid As Long, trans As Boolean, failure As Long, message As String
    On Error GoTo Failed
    pid = RouteProfile()
    If outcome <> "Accepted" And outcome <> "Returned" And outcome <> "Cancelled" Then Err.Raise vbObjectError + 1809, , "Choose an available review outcome."
    If outcome = "Returned" And Len(Trim$(Replace(Replace(Replace(notes, vbCr, ""), vbLf, ""), vbTab, ""))) = 0 Then Err.Raise vbObjectError + 1810, , "Explain why the review is returned."
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    Set r = db.OpenRecordset("SELECT * FROM ReviewTask WHERE TaskID=" & taskID, dbOpenDynaset)
    If r.EOF Then Err.Raise vbObjectError + 1811, , "Select a review assignment."
    If r!TaskState <> "Pending" Then Err.Raise vbObjectError + 1812, , "This review already has an outcome."
    If outcome = "Cancelled" Then
        If pid <> r!RequestedBy Then Err.Raise vbObjectError + 1813, , "Switch to the requesting contributor to cancel this assignment."
    Else
        If pid <> r!AssignedTo Then Err.Raise vbObjectError + 1814, , "Switch to the assigned reviewer to record an outcome."
        If Nz(RouteValue(db, "SELECT IsArchived FROM Report WHERE ReportID=" & r!ReportID), True) Then Err.Raise vbObjectError + 1815, , "Archived reports cannot receive review outcomes."
        If Nz(RouteValue(db, "SELECT CurrentRevisionID FROM Report WHERE ReportID=" & r!ReportID), 0) <> r!RevisionID Then Err.Raise vbObjectError + 1816, , "This review is stale. The requester must cancel it and assign the latest saved revision."
    End If
    r.Edit: r!TaskState = outcome: r!FinishedAt = Now(): r.Update: r.Close: Set r = Nothing
    Call RouteEvent(db, taskID, pid, outcome, notes)
    w.CommitTrans: trans = False: FinishReview = True: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not r Is Nothing Then r.Close: Set r = Nothing
    If trans Then w.Rollback
    On Error GoTo 0
    Err.Raise failure, "FinishReview", message
End Function
Public Function OpenMyWork(Optional ByVal reportID As Long = 0) As Boolean
    DoCmd.OpenForm "frmMyWork"
    Forms!frmMyWork!txtReportID = reportID
    Forms!frmMyWork!txtExpectedRevision = Nz(DLookup("CurrentRevisionID", "Report", "ReportID=" & reportID), 0)
    Forms!frmMyWork!lblAssignment.Caption = "Assign saved report " & reportID & " / revision ID " & Forms!frmMyWork!txtExpectedRevision & ". Open from a saved draft to assign it."
    Call LoadMyWork(): OpenMyWork = True
End Function
Public Function RoutingFromEditor() As Boolean
    RoutingFromEditor = OpenMyWork(Nz(Forms!frmEditor!txtReportID, 0))
End Function
Public Function LoadMyWork() As Boolean
    Dim sql As String, pid As Long, reportID As Long
    pid = RouteProfile()
    reportID = Nz(Forms!frmMyWork!txtReportID, 0)
    Forms!frmMyWork!txtExpectedRevision = Nz(RouteScalar("SELECT CurrentRevisionID FROM Report WHERE ReportID=" & reportID), 0)
    Forms!frmMyWork!lblAssignment.Caption = "Assign saved report " & reportID & " / revision ID " & Forms!frmMyWork!txtExpectedRevision & ". Open from a saved draft to assign it."
    Forms!frmMyWork!lblProfile.Caption = "Active contributor: " & DLookup("DisplayName", "UserProfile", "UserProfileID=" & pid) & ". Profiles are attribution, not authentication."
    Forms!frmMyWork!cboContributor.RowSource = "SELECT UserProfileID,DisplayName FROM UserProfile ORDER BY DisplayName": Forms!frmMyWork!cboContributor.Requery
    Forms!frmMyWork!cboReviewer.RowSource = Forms!frmMyWork!cboContributor.RowSource: Forms!frmMyWork!cboReviewer.Requery
    sql = "SELECT t.TaskID,r.MemberName & ' | ' & r.FormKind & ' | ' & Nz(Format(r.PeriodEnd,'yyyy-mm-dd'),'No period') AS [Member and report],p.DisplayName AS Reviewer,t.TaskState AS [Review step],t.DueDate AS [Due date],IIf(t.RevisionID=r.CurrentRevisionID,'Current','Stale') AS [Saved version],r.IsArchived FROM (ReviewTask AS t INNER JOIN Report AS r ON t.ReportID=r.ReportID) INNER JOIN UserProfile AS p ON t.AssignedTo=p.UserProfileID"
    If Not Nz(Forms!frmMyWork!chkAll, False) Then sql = sql & " WHERE t.AssignedTo=" & pid & " AND t.TaskState='Pending'"
    Forms!frmMyWork!lstTasks.RowSource = sql & " ORDER BY t.TaskID DESC": Forms!frmMyWork!lstTasks.Requery
    Forms!frmMyWork!txtHistory = Null: LoadMyWork = True
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
Public Function ContributorUI(ByVal action As String) As Boolean
    On Error GoTo Failed
    If action = "Add" Then
        Call AddContributor(Nz(Forms!frmMyWork!txtContributor, ""))
        Forms!frmMyWork!txtContributor = Null
    Else
        Call SwitchContributor(Nz(Forms!frmMyWork!cboContributor, 0))
    End If
    Call LoadMyWork(): ContributorUI = True: Exit Function
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
    Dim id As Long
    On Error GoTo Failed
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 1817, , "Close the current editor before opening another report."
    id = Nz(DLookup("ReportID", "ReviewTask", "TaskID=" & Nz(Forms!frmMyWork!lstTasks, 0)), 0)
    If id = 0 Then Err.Raise vbObjectError + 1811, , "Select a review assignment."
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id): OpenReviewTaskDraft = True: Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Review draft"
End Function
