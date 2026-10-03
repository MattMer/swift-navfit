Attribute VB_Name = "RoutingHarness"
Option Explicit
Private checks As Long
Private lastFailure As String
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 1850, , message
    checks = checks + 1
End Sub
Private Function Denied(ByVal action As String, ByVal id As Long, Optional ByVal argument As Long = 0, Optional ByVal text As String = "") As Boolean
    On Error GoTo Failed
    Select Case action
        Case "assign": Call AssignReview(id, argument, CLng(text), Null, "Instruction")
        Case "finish": Call FinishReview(id, "Accepted", text)
        Case "return": Call FinishReview(id, "Returned", text)
        Case "cancel": Call FinishReview(id, "Cancelled", text)
        Case "switch": Call SwitchContributor(id)
        Case "profile": Call AddContributor(text)
    End Select
    Exit Function
Failed:
    lastFailure = Err.Number & " " & Err.Description
    Denied = True
End Function
Private Function CountRows(ByVal table As String) As Long
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT Count(*) FROM [" & table & "]", dbOpenSnapshot)
    CountRows = r.Fields(0).Value: r.Close
End Function
Private Sub EventRule(ByVal rule As String)
    Dim d As DAO.Database
    Set d = CurrentDb
    d.TableDefs("ReviewTaskEvent").Fields("EventKind").ValidationRule = rule
End Sub
Public Function TestRouting() As String
    Dim requester As Long, reviewer As Long, reportID As Long, revision As Long, task As Long, newer As Long
    Dim before As Long, eventCount As Long, body As String, r As DAO.Recordset, failure As Long, message As String
    On Error GoTo Failed
    Call ConfigureWorkspace("SYNTHETIC ROUTING", "Requester")
    requester = DLookup("ActiveProfileID", "Workspace")
    reviewer = AddContributor("Reviewer O'Name")
    Check reviewer <> requester, "Contributor identity reused"
    Check Denied("profile", 0, 0, "Reviewer O'Name"), "Duplicate contributor accepted"
    Check Denied("profile", 0, 0, " "), "Blank contributor accepted"
    reportID = CommitDraft(0, "EVAL", "Review me", "Synthetic Member", Null, Null, "", "Initial", "Official narrative")
    revision = DLookup("CurrentRevisionID", "Report", "ReportID=" & reportID)
    before = DCount("*", "ReportRevision")
    Check Denied("assign", reportID, revision, "99999"), "Missing reviewer accepted"
    Check Denied("assign", reportID, 99999, CStr(reviewer)), "Wrong revision accepted"
    body = String$(8000, "Q") & vbCrLf & "Exact instructions'End"
    EventRule "<>'Assigned'"
    Check Denied("assign", reportID, revision, CStr(reviewer)), "Injected assignment failure accepted"
    EventRule ""
    Check CountRows("ReviewTask") = 0 And CountRows("ReviewTaskEvent") = 0, "Failed assignment partly committed: tasks=" & CountRows("ReviewTask") & " events=" & CountRows("ReviewTaskEvent") & " rejected=" & lastFailure
    task = AssignReview(reportID, revision, reviewer, Date - 1, body)
    Check task > 0 And CountRows("ReviewTaskEvent") = 1, "Assignment missing: task=" & task & " events=" & CountRows("ReviewTaskEvent")
    Set r = CurrentDb.OpenRecordset("SELECT * FROM ReviewTask WHERE TaskID=" & task, dbOpenSnapshot)
    Check StrComp(r!Instructions, body, vbBinaryCompare) = 0 And r!DueDate = Date - 1, "Instructions or due date changed": r.Close
    Check Denied("assign", reportID, revision, CStr(reviewer)), "Parallel pending assignment accepted"
    Check Denied("finish", task), "Requester accepted reviewer task"
    Check SwitchContributor(reviewer), "Switch failed"
    Check OpenMyWork(), "My Work failed"
    Check Forms!frmMyWork!lstTasks.ListCount = 1, "Reviewer pending filter wrong"
    Forms!frmMyWork!lstTasks = task: Check TaskDetails(), "Task history failed"
    Check InStr(Forms!frmMyWork!txtHistory, body) > 0 And InStr(Forms!frmMyWork!txtHistory, "overdue") > 0, "History or overdue missing"
    Check Denied("cancel", task), "Reviewer cancelled requester task"
    Check Denied("return", task, 0, " "), "Return without reason accepted"
    EventRule "<>'Accepted'"
    Check Denied("finish", task), "Injected outcome failure accepted"
    EventRule ""
    Check DLookup("TaskState", "ReviewTask", "TaskID=" & task) = "Pending" And CountRows("ReviewTaskEvent") = 1, "Failed outcome partly committed"
    Check FinishReview(task, "Returned", "Needs revision"), "Return failed"
    Check Denied("finish", task), "Finished review overwritten"
    Call LoadMyWork(): Check Forms!frmMyWork!lstTasks.ListCount = 0, "Returned review still pending"
    Call SwitchContributor(requester)
    task = AssignReview(reportID, revision, reviewer, Null, "Second pass")
    Call CommitDraft(reportID, "EVAL", "Later saved", "Synthetic Member", Null, Null, "", "Later", "Later narrative")
    newer = DLookup("CurrentRevisionID", "Report", "ReportID=" & reportID)
    Call SwitchContributor(reviewer)
    Check Denied("finish", task), "Stale review accepted"
    Call LoadMyWork(): Forms!frmMyWork!lstTasks = task: Call TaskDetails()
    Check InStr(Forms!frmMyWork!txtHistory, "STALE") > 0, "Stale warning absent"
    Call SwitchContributor(requester): Check FinishReview(task, "Cancelled", "Superseded"), "Stale cancellation failed"
    task = AssignReview(reportID, newer, reviewer, Null, "Latest")
    Call SwitchContributor(reviewer)
    Forms!frmMyWork!chkAll = True: Call LoadMyWork(): Forms!frmMyWork!lstTasks = task
    Forms!frmMyWork!txtNotes = "Accepted internal review"
    Check FinishReviewUI("Accepted"), "UI accept failed"
    Check DLookup("TaskState", "ReviewTask", "TaskID=" & task) = "Accepted", "Accepted state missing"
    Check DLookup("WorkflowState", "Report", "ReportID=" & reportID) = "Draft", "Internal outcome changed draft state"
    Call SwitchContributor(requester)
    Check OpenMyWork(reportID), "Assignment context failed"
    Forms!frmMyWork!cboReviewer = reviewer: Forms!frmMyWork!txtNotes = "UI assignment"
    Check AssignReviewUI(), "UI assignment failed"
    task = DMax("TaskID", "ReviewTask")
    Check DCount("*", "ReportRevision") = before + 1, "Routing changed revisions"
    DoCmd.OpenForm "frmEditor", , , , , , CStr(reportID)
    Forms!frmEditor!txtTitle = "Pending text"
    Check Denied("switch", reviewer), "Switched with open editor"
    Check Forms!frmEditor!txtTitle = "Pending text", "Routing changed pending editor"
    Check RoutingFromEditor(), "Editor routing action failed"
    DoCmd.Close acForm, "frmEditor", acSaveNo
    Call ArchiveDraft(reportID, True): Call SwitchContributor(reviewer)
    Check Denied("finish", task), "Archived review accepted"
    Call SwitchContributor(requester): Check FinishReview(task, "Cancelled", "Archived"), "Archived cancellation blocked"
    Check Denied("assign", reportID, newer, CStr(reviewer)), "Archived report assigned"
    DoCmd.Close acForm, "frmMyWork", acSaveNo
    TestRouting = "PASS: " & checks & " routing assertions": Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    EventRule ""
    TestRouting = "FAIL: " & failure & " " & message
End Function
