Attribute VB_Name = "modRouteTemplates"
Option Compare Binary
Option Explicit
'F01 review chain: a routed report is one shared draft. The reviewer at the
'current step holds it and may save; Forward passes it up, Send back and Pull back
'return it to an earlier step, and the chain then runs upward again from there.
Private gridContext As Boolean, saveNotice As String
Public Function CreateRouteTemplate(ByVal title As String, ByVal assigneesCSV As String, ByVal instructions As String) As Long
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim ids As Object, key As Variant, r As DAO.Recordset, s As DAO.Recordset, id As Long, ordinal As Long, w As DAO.Workspace, failure As Long, message As String, operation As Long, author As Long
    author = MajorAuthor()
    If Len(Trim$(title)) = 0 Or Len(title) > 120 Then Err.Raise vbObjectError + 4300, "modRouteTemplates.CreateRouteTemplate", "Enter a route title."
    If Len(Trim$(assigneesCSV)) = 0 Then Err.Raise vbObjectError + 4300, "modRouteTemplates.CreateRouteTemplate", "Choose at least one reviewer."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = ownedDB.OpenRecordset("RouteTemplate", dbOpenDynaset)
    r.AddNew: r!Title = title: r!IsArchived = False: r.Update: r.Bookmark = r.LastModified: id = r!TemplateID: r.Close
    Set s = ownedDB.OpenRecordset("RouteTemplateStep", dbOpenDynaset)
    For Each key In Split(assigneesCSV, ",")
        If DCount("*", "UserProfile", "IsRetired=False AND UserProfileID=" & CLng(Trim$(key))) = 0 Then Err.Raise vbObjectError + 4300, "modRouteTemplates.CreateRouteTemplate", "A chosen reviewer is not an active contributor."
        Set ids = MajorIDs(CStr(key))
        ordinal = ordinal + 1
        s.AddNew: s!TemplateID = id: s!StepOrdinal = ordinal: s!AssigneeID = CLng(Trim$(key)): s!Instructions = instructions: s.Update
    Next
    s.Close
    operation = BeginOperation("Create route template", "Template " & id & ": " & title & "; ordered contributors: " & assigneesCSV & "; instructions: " & instructions)
    FinishOperation operation, "Completed", "Template stored; future routes copy its steps."
    w.CommitTrans: CreateRouteTemplate = id: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRouteTemplates", "CreateRouteTemplate", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    If Not s Is Nothing Then SwiftClose s, swiftPrimary, "s"
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
' Deleting a template hides it from new reviews. Routes already started keep their
' own copied reviewers in order, so reports under review are unaffected.
Public Sub DeleteRouteTemplate(ByVal templateID As Long)
    Dim operation As Long, title As String
    title = Nz(MajorScalar("SELECT Title FROM RouteTemplate WHERE TemplateID=" & templateID & " AND IsArchived=False"), "")
    If Len(title) = 0 Then Err.Raise vbObjectError + 4321, "modRouteTemplates.DeleteRouteTemplate", "Choose a route template to delete."
    operation = BeginOperation("Delete route template", "Template " & templateID & ": " & title)
    CurrentDb.Execute "UPDATE RouteTemplate SET IsArchived=True WHERE TemplateID=" & templateID, dbFailOnError
    FinishOperation operation, "Completed", "Template hidden from new reviews; existing routes keep their reviewers."
End Sub
' Swap a report's active route for a new one from another template. The old route
' is kept as cancelled history; the new one starts at its step 1.
Public Function ReplaceRoute(ByVal runID As Long, ByVal templateID As Long) As Long
    Dim db As DAO.Database, w As DAO.Workspace, r As DAO.Recordset, reportID As Long, groupID As Long
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = db.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & runID, dbOpenDynaset)
    If r.EOF Then Err.Raise vbObjectError + 4308, "modRouteTemplates.ReplaceRoute", "Route does not exist."
    If r!RunState <> "Pending" And r!RunState <> "Accepted" Then Err.Raise vbObjectError + 4309, "modRouteTemplates.ReplaceRoute", "This route is finished. Start a new route instead."
    reportID = r!ReportID: groupID = Nz(r!GroupID, 0)
    If groupID > 0 Then
        If Nz(MajorScalar("SELECT IsArchived FROM SummaryGroup WHERE GroupID=" & groupID), True) Then groupID = 0
    End If
    r.Edit: r!RunState = "Cancelled": r.Update: r.Close: Set r = Nothing
    RouteEventAdd db, runID, "Replaced", Null, Null, CurrentRevision(reportID), "Replaced with route template: " & Nz(MajorScalar("SELECT Title FROM RouteTemplate WHERE TemplateID=" & templateID), "")
    ReplaceRoute = StartRoute(reportID, templateID, groupID, False)
    w.CommitTrans: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRouteTemplates", "ReplaceRoute", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
Public Function StartRoute(ByVal reportID As Long, ByVal templateID As Long, Optional ByVal groupID As Long = 0, Optional ByVal ownTransaction As Boolean = True) As Long
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim w As DAO.Workspace, r As DAO.Recordset, s As DAO.Recordset, t As DAO.Recordset, id As Long, failure As Long, message As String
    Set w = DBEngine.Workspaces(0): If ownTransaction Then w.BeginTrans
    On Error GoTo Failed
    If Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & reportID), True) Then Err.Raise vbObjectError + 4301, "modRouteTemplates.StartRoute", "Select an active report."
    If Nz(MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & reportID), "") <> "Draft" Then Err.Raise vbObjectError + 4301, "modRouteTemplates.StartRoute", "Reopen the approved report before starting another route."
    If Nz(MajorScalar("SELECT Count(*) FROM RouteRun WHERE ReportID=" & reportID & " AND RunState='Pending'"), 0) > 0 Then Err.Raise vbObjectError + 4302, "modRouteTemplates.StartRoute", "Finish or cancel the active route first."
    If Nz(MajorScalar("SELECT IsArchived FROM RouteTemplate WHERE TemplateID=" & templateID), True) Then Err.Raise vbObjectError + 4303, "modRouteTemplates.StartRoute", "Select an active template."
    Set t = ownedDB.OpenRecordset("SELECT * FROM RouteTemplateStep WHERE TemplateID=" & templateID & " ORDER BY StepOrdinal", dbOpenSnapshot)
    If t.EOF Then Err.Raise vbObjectError + 4304, "modRouteTemplates.StartRoute", "Route template is empty."
    Set r = ownedDB.OpenRecordset("RouteRun", dbOpenDynaset)
    r.AddNew: r!ReportID = reportID: r!RevisionID = CurrentRevision(reportID): r!TemplateID = templateID: r!CurrentStep = 1: r!RunState = "Pending": r!RequestedBy = MajorAuthor(): r!RequestedAt = Now()
    If groupID > 0 Then
        If Nz(MajorScalar("SELECT IsArchived FROM SummaryGroup WHERE GroupID=" & groupID), True) Then Err.Raise vbObjectError + 4305, "modRouteTemplates.StartRoute", "Recover the archived group before starting review."
        If DCount("*", "GroupMember", "GroupID=" & groupID & " AND ReportID=" & reportID) <> 1 Then Err.Raise vbObjectError + 4305, "modRouteTemplates.StartRoute", "Report is not in the selected group."
        r!GroupID = groupID: r!GroupStamp = GroupRouteStamp(groupID)
    End If
    r.Update: r.Bookmark = r.LastModified: id = r!RunID: r.Close
    Set s = ownedDB.OpenRecordset("RouteRunStep", dbOpenDynaset)
    Do While Not t.EOF
        s.AddNew: s!RunID = id: s!StepOrdinal = t!StepOrdinal: s!AssigneeID = t!AssigneeID: s!Instructions = t!Instructions: s!StepState = "Pending"
        If t!StepOrdinal = 1 Then s!IncomingRevisionID = CurrentRevision(reportID)
        s.Update: t.MoveNext
    Loop
    s.Close: t.Close
    RouteEventAdd ownedDB, id, "Start", 0, 1, CurrentRevision(reportID), ""
    If ownTransaction Then w.CommitTrans
    StartRoute = id: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRouteTemplates", "StartRoute", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    If ownTransaction Then SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    If Not s Is Nothing Then SwiftClose s, swiftPrimary, "s"
    If Not t Is Nothing Then SwiftClose t, swiftPrimary, "t"
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function
' A route stays current while its report is active and, for a group route, the
' group's membership is unchanged. Saves no longer make it stale: the holder edits
' the shared draft, and other contributors are refused at save.
Public Function RouteIsCurrent(ByVal runID As Long) As Boolean
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset
    Set r = ownedDB.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & runID, dbOpenSnapshot)
    If r.EOF Then r.Close: Exit Function
    RouteIsCurrent = Not Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & r!ReportID), True)
    If Not IsNull(r!GroupID) Then
        If Not RouteGroupCurrent(r!GroupStamp, r!GroupID) Then RouteIsCurrent = False
    End If
    r.Close
End Function
'Accepted forwards the shared draft to the next step; Returned sends it back one
'step (to the requester from step 1).
Public Sub DecideRouteStep(ByVal runID As Long, ByVal decision As String, ByVal note As String)
    If decision = "Accepted" Then
        ForwardRoute runID, note
    ElseIf decision = "Returned" Then
        SendBackRoute runID, Nz(MajorScalar("SELECT CurrentStep FROM RouteRun WHERE RunID=" & runID), 0) - 1, note
    Else
        Err.Raise vbObjectError + 4306, "modRouteTemplates.DecideRouteStep", "Choose Accepted or Returned."
    End If
End Sub
'A signed report is finished: its open route (pending, or awaiting local approval)
'and pending assignments end, with a history entry saying why. The caller owns the
'signing transaction. Returns the number of reviews ended.
Public Function EndReviewsForSigning(ByVal reportID As Long) As Long
    Dim db As DAO.Database, r As DAO.Recordset, note As String, revision As Long
    Set db = CurrentDb: revision = CurrentRevision(reportID)
    note = "Ended because the report was marked signed."
    Set r = db.OpenRecordset("SELECT * FROM RouteRun WHERE ReportID=" & reportID & " AND RunState In ('Pending','Accepted')", dbOpenDynaset)
    Do While Not r.EOF
        r.Edit: r!RunState = "Cancelled": r.Update
        RouteEventAdd db, r!RunID, "Cancelled", r!CurrentStep, r!CurrentStep, Null, note
        EndReviewsForSigning = EndReviewsForSigning + 1: r.MoveNext
    Loop
    r.Close
    Set r = db.OpenRecordset("SELECT * FROM ReviewTask WHERE ReportID=" & reportID & " AND TaskState='Pending'", dbOpenDynaset)
    Do While Not r.EOF
        r.Edit: r!TaskState = "Cancelled": r!FinishedAt = Now(): r!FinishedRevisionID = revision: r.Update
        db.Execute "INSERT INTO ReviewTaskEvent (TaskID,AuthorProfileID,EventAt,EventKind,Notes) VALUES (" & r!TaskID & "," & MajorAuthor() & ",Now(),'Cancelled'," & QuoteSQL(note) & ")", dbFailOnError
        EndReviewsForSigning = EndReviewsForSigning + 1: r.MoveNext
    Loop
    r.Close
End Function
Public Sub CancelRoute(ByVal runID As Long)
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    If Nz(MajorScalar("SELECT RequestedBy FROM RouteRun WHERE RunID=" & runID), 0) <> MajorAuthor() Then Err.Raise vbObjectError + 4312, "modRouteTemplates.CancelRoute", "Only the requester can cancel this route."
    ownedDB.Execute "UPDATE RouteRun SET RunState='Cancelled' WHERE RunID=" & runID & " AND RunState='Pending'", dbFailOnError
End Sub
Public Function ApprovalIsCurrent(ByVal runID As Long) As Boolean
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    If Nz(MajorScalar("SELECT RunState FROM RouteRun WHERE RunID=" & runID), "") <> "Approved" Then Exit Function
    ApprovalIsCurrent = RouteIsCurrent(runID)
End Function
Public Sub ApproveSavedRoute(ByVal runID As Long, ByVal note As String)
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset, e As DAO.Recordset, w As DAO.Workspace, operation As Long, failure As Long, message As String
    If Len(Trim$(note)) = 0 Then Err.Raise vbObjectError + 4313, "modRouteTemplates.ApproveSavedRoute", "Record a local approval note."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = ownedDB.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & runID, dbOpenDynaset)
    If r.EOF Then Err.Raise vbObjectError + 4308, "modRouteTemplates.ApproveSavedRoute", "Route does not exist."
    If r!RunState <> "Accepted" Or Not RouteIsCurrent(runID) Then Err.Raise vbObjectError + 4314, "modRouteTemplates.ApproveSavedRoute", "Complete a current route first."
    If Nz(r!ApprovedBy, 0) <> MajorAuthor() Then Err.Raise vbObjectError + 4315, "modRouteTemplates.ApproveSavedRoute", "The final assigned reviewer must record local approval."
    If Nz(MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & r!ReportID), "") <> "Draft" Then Err.Raise vbObjectError + 4316, "modRouteTemplates.ApproveSavedRoute", "Report is already locked."
    operation = BeginOperation("Local approval", "Route " & runID & ": " & note)
    RetainReportCheckpoint r!ReportID, "Local approval"
    ownedDB.Execute "UPDATE Report SET WorkflowState='Approved' WHERE ReportID=" & r!ReportID, dbFailOnError
    Set e = ownedDB.OpenRecordset("ReportEvent", dbOpenDynaset)
    e.AddNew: e!ReportID = r!ReportID: e!AuthorProfileID = MajorAuthor(): e!EventAt = Now(): e!EventKind = "Local approval": e.Update: e.Bookmark = e.LastModified: AuditRecord "Report event", CLng(e!ReportEventID), operation: e.Close
    r.Edit: r!RunState = "Approved": r!ApprovedAt = Now(): r!ApprovedRevisionID = CurrentRevision(r!ReportID): r.Update: r.Close
    FinishOperation operation, "Completed", "Saved revision locally approved and editing locked; no official signature or submission."
    w.CommitTrans: Exit Sub
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRouteTemplates", "ApproveSavedRoute", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    If Not e Is Nothing Then SwiftClose e, swiftPrimary, "e"
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Sub
Public Sub ReopenApprovedReport(ByVal reportID As Long, ByVal note As String)
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim w As DAO.Workspace, e As DAO.Recordset, operation As Long, failure As Long, message As String
    If Len(Trim$(note)) = 0 Then Err.Raise vbObjectError + 4317, "modRouteTemplates.ReopenApprovedReport", "Record a reopening reason."
    Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    If Nz(MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & reportID), "") <> "Approved" Then Err.Raise vbObjectError + 4318, "modRouteTemplates.ReopenApprovedReport", "Select a locally approved report."
    If Nz(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & reportID), True) Then Err.Raise vbObjectError + 4318, "modRouteTemplates.ReopenApprovedReport", "Recover the archived report first."
    operation = BeginOperation("Reopen approval", "Report " & reportID & ": " & note)
    ownedDB.Execute "UPDATE RouteRun SET RunState='Reopened' WHERE ReportID=" & reportID & " AND RunState='Approved'", dbFailOnError
    ownedDB.Execute "UPDATE Report SET WorkflowState='Draft' WHERE ReportID=" & reportID, dbFailOnError
    Set e = ownedDB.OpenRecordset("ReportEvent", dbOpenDynaset)
    e.AddNew: e!ReportID = reportID: e!AuthorProfileID = MajorAuthor(): e!EventAt = Now(): e!EventKind = "Approval reopened": e.Update: e.Bookmark = e.LastModified: AuditRecord "Report event", CLng(e!ReportEventID), operation: e.Close
    FinishOperation operation, "Completed", "Editing reopened; previous local approval retained as history."
    w.CommitTrans: Exit Sub
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRouteTemplates", "ReopenApprovedReport", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary
    If Not e Is Nothing Then SwiftClose e, swiftPrimary, "e"
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Sub

Private Sub RouteEventAdd(ByVal db As DAO.Database, ByVal runID As Long, ByVal kind As String, ByVal fromStep As Variant, ByVal toStep As Variant, ByVal revisionID As Variant, ByVal notes As String)
    Dim e As DAO.Recordset
    If Not IsNull(revisionID) Then
        If kind <> "Grid change" And kind <> "Higher change" And kind <> "Override" Then
            RetainReportCheckpoint CLng(MajorScalar("SELECT ReportID FROM ReportRevision WHERE RevisionID=" & CLng(revisionID))), kind
        ElseIf RevisionRetentionKind(CLng(revisionID)) = "Working" Then
            db.Execute "UPDATE RevisionRetention SET RetentionKind='Protected' WHERE RevisionID=" & CLng(revisionID), dbFailOnError
        End If
    End If
    Set e = db.OpenRecordset("RouteRunEvent", dbOpenDynaset, dbAppendOnly)
    e.AddNew: e!RunID = runID: e!EventKind = kind: e!FromStep = fromStep: e!ToStep = toStep: e!RevisionID = revisionID
    e!AuthorProfileID = MajorAuthor(): e!EventAt = Now()
    If Len(notes) > 0 Then e!Notes = notes
    e.Update: e.Close
End Sub
'The run, refused when it is finished or its report or group can no longer take review.
Private Function RouteForAction(ByVal db As DAO.Database, ByVal runID As Long, ByVal procedure As String) As DAO.Recordset
    Dim r As DAO.Recordset
    Set r = db.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & runID, dbOpenDynaset)
    If r.EOF Then Err.Raise vbObjectError + 4308, "modRouteTemplates." & procedure, "Route does not exist."
    If r!RunState <> "Pending" And r!RunState <> "Accepted" Then Err.Raise vbObjectError + 4309, "modRouteTemplates." & procedure, "This route is finished. Start a new route if needed."
    If Not RouteIsCurrent(runID) Then Err.Raise vbObjectError + 4309, "modRouteTemplates." & procedure, "The group's membership changed or the report is archived. The requester can restart the route."
    If Nz(MajorScalar("SELECT WorkflowState FROM Report WHERE ReportID=" & r!ReportID), "") <> "Draft" Then Err.Raise vbObjectError + 4316, "modRouteTemplates." & procedure, "Reopen the report before changing its review."
    Set RouteForAction = r
End Function
Private Function RouteStepCount(ByVal runID As Long) As Long
    RouteStepCount = Nz(MajorScalar("SELECT Count(*) FROM RouteRunStep WHERE RunID=" & runID), 0)
End Function
'Who holds a run: the current step's reviewer, the requester at step 0, or the
'final reviewer once every step has forwarded.
Private Function RouteRunHolder(ByVal r As DAO.Recordset) As Long
    If r!RunState = "Accepted" Then
        RouteRunHolder = Nz(r!ApprovedBy, 0)
    ElseIf r!CurrentStep = 0 Then
        RouteRunHolder = r!RequestedBy
    Else
        RouteRunHolder = Nz(MajorScalar("SELECT AssigneeID FROM RouteRunStep WHERE RunID=" & r!RunID & " AND StepOrdinal=" & r!CurrentStep), 0)
    End If
End Function
'A deleted contributor leaves every route template and active route. Later steps
'move up; when the holder is removed the next step holds the draft; a route with no
'reviewer left, or still with a deleted requester at step 0, is cancelled. Runs in
'the caller's transaction. Returns a summary.
Public Function RouteRemoveContributor(ByVal db As DAO.Database, ByVal profileID As Long, ByVal name As String) As String
    Dim ids As Collection, id As Variant, r As DAO.Recordset, n As Long, before As Long, k As Long, revision As Long, templates As Long, emptied As Long, runs As Long, cancelled As Long, note As String
    note = name & " was deleted as a contributor."
    Set ids = RouteIDs(db, "SELECT DISTINCT TemplateID FROM RouteTemplateStep WHERE AssigneeID=" & profileID)
    For Each id In ids
        n = RouteRenumber(db, "SELECT * FROM RouteTemplateStep WHERE TemplateID=" & id & " ORDER BY StepOrdinal", profileID, 0, before)
        If n = 0 Then
            Set r = db.OpenRecordset("SELECT * FROM RouteTemplate WHERE TemplateID=" & id, dbOpenDynaset)
            r.Edit: r!IsArchived = True: r.Update: r.Close: emptied = emptied + 1
        End If
        templates = templates + 1
    Next
    Set ids = RouteIDs(db, "SELECT DISTINCT s.RunID FROM RouteRunStep AS s INNER JOIN RouteRun AS u ON s.RunID=u.RunID WHERE s.AssigneeID=" & profileID & " AND u.RunState In ('Pending','Accepted')")
    For Each id In ids
        Set r = db.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & id, dbOpenDynaset)
        k = r!CurrentStep: revision = CurrentRevision(r!ReportID)
        n = RouteRenumber(db, "SELECT * FROM RouteRunStep WHERE RunID=" & id & " ORDER BY StepOrdinal", profileID, k, before)
        r.Edit
        If n = 0 Then
            r!RunState = "Cancelled": cancelled = cancelled + 1
        ElseIf r!RunState = "Accepted" Then
            r!CurrentStep = n
            If Nz(r!ApprovedBy, 0) = profileID Then r!ApprovedBy = RouteStepAssignee(db, CLng(id), n)
        ElseIf k - before > n Then
            'The removed holder was the last step; everyone left has forwarded.
            r!CurrentStep = n: r!RunState = "Accepted": r!ApprovedBy = RouteStepAssignee(db, CLng(id), n): r!ApprovedAt = Now()
        Else
            r!CurrentStep = k - before
            If k - before >= 1 Then RouteSetIncoming db, CLng(id), k - before, revision
        End If
        r.Update: r.Bookmark = r.LastModified
        RouteEventAdd db, CLng(id), IIf(n = 0, "Cancelled", "Reviewer removed"), k, r!CurrentStep, revision, note & IIf(n = 0, " No reviewer remains; the route is cancelled.", "")
        r.Close: runs = runs + 1
    Next
    Set r = db.OpenRecordset("SELECT * FROM RouteRun WHERE RequestedBy=" & profileID & " AND RunState='Pending' AND CurrentStep=0", dbOpenDynaset)
    Do While Not r.EOF
        r.Edit: r!RunState = "Cancelled": r.Update
        RouteEventAdd db, r!RunID, "Cancelled", 0, 0, CurrentRevision(r!ReportID), note & " The report was back with them as requester; the route is cancelled."
        cancelled = cancelled + 1: r.MoveNext
    Loop
    r.Close
    RouteRemoveContributor = "Removed from " & templates & " route template(s) (" & emptied & " left empty and deleted) and " & runs & " active route(s); " & cancelled & " route(s) cancelled;"
End Function
Private Function RouteIDs(ByVal db As DAO.Database, ByVal sql As String) As Collection
    Dim r As DAO.Recordset
    Set RouteIDs = New Collection
    Set r = db.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not r.EOF: RouteIDs.Add CLng(r.Fields(0).Value): r.MoveNext: Loop
    r.Close
End Function
'Deletes a contributor's steps and closes the gaps. Returns the steps left; before
'counts removed steps below currentStep.
Private Function RouteRenumber(ByVal db As DAO.Database, ByVal sql As String, ByVal profileID As Long, ByVal currentStep As Long, ByRef before As Long) As Long
    Dim s As DAO.Recordset
    before = 0
    Set s = db.OpenRecordset(sql, dbOpenDynaset)
    Do While Not s.EOF
        If s!AssigneeID = profileID Then
            If s!StepOrdinal < currentStep Then before = before + 1
            s.Delete
        Else
            RouteRenumber = RouteRenumber + 1
            If s!StepOrdinal <> RouteRenumber Then s.Edit: s!StepOrdinal = RouteRenumber: s.Update
        End If
        s.MoveNext
    Loop
    s.Close
End Function
Private Function RouteStepAssignee(ByVal db As DAO.Database, ByVal runID As Long, ByVal ordinal As Long) As Long
    Dim s As DAO.Recordset
    Set s = db.OpenRecordset("SELECT AssigneeID FROM RouteRunStep WHERE RunID=" & runID & " AND StepOrdinal=" & ordinal, dbOpenSnapshot)
    If Not s.EOF Then RouteStepAssignee = s!AssigneeID
    s.Close
End Function
Private Sub RouteSetIncoming(ByVal db As DAO.Database, ByVal runID As Long, ByVal ordinal As Long, ByVal revision As Long)
    Dim s As DAO.Recordset
    Set s = db.OpenRecordset("SELECT * FROM RouteRunStep WHERE RunID=" & runID & " AND StepOrdinal=" & ordinal, dbOpenDynaset)
    If Not s.EOF Then
        If IsNull(s!IncomingRevisionID) Then s.Edit: s!IncomingRevisionID = revision: s.Update
    End If
    s.Close
End Sub
Public Function RouteHolder(ByVal runID As Long) As Long
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & runID, dbOpenSnapshot)
    If Not r.EOF Then
        If r!RunState = "Pending" Or r!RunState = "Accepted" Then RouteHolder = RouteRunHolder(r)
    End If
    r.Close
End Function
'Steps from toStep upward run again; their earlier decisions stay in RouteRunEvent.
Private Sub RouteReturnTo(ByVal db As DAO.Database, ByVal r As DAO.Recordset, ByVal toStep As Long, ByVal kind As String, ByVal note As String)
    Dim fromStep As Long, revision As Long, runID As Long
    Dim s As DAO.Recordset
    fromStep = r!CurrentStep: revision = CurrentRevision(r!ReportID): runID = r!RunID
    'Route rows change through recordsets only: a SQL update inside the same
    'transaction can collide with the recordset's own lock (error 3218).
    Set s = db.OpenRecordset("SELECT * FROM RouteRunStep WHERE RunID=" & runID & " AND StepOrdinal>=" & IIf(toStep < 1, 1, toStep), dbOpenDynaset)
    Do While Not s.EOF
        s.Edit: s!StepState = "Pending": s!DecisionNote = Null: s!DecisionAt = Null: s!OutgoingRevisionID = Null
        If s!StepOrdinal = toStep Then s!IncomingRevisionID = revision Else s!IncomingRevisionID = Null
        s.Update: s.MoveNext
    Loop
    s.Close
    r.Edit: r!CurrentStep = toStep: r!RunState = "Pending": r!ApprovedBy = Null: r!ApprovedAt = Null: r.Update
    RouteEventAdd db, runID, kind, fromStep, toStep, revision, note
End Sub
'The holder passes the shared draft, as now saved, to the next step. At step 0 the
'requester resubmits to step 1. The last step's Forward completes the route.
Public Sub ForwardRoute(ByVal runID As Long, ByVal note As String)
    Dim db As DAO.Database, w As DAO.Workspace, r As DAO.Recordset, s As DAO.Recordset, k As Long, n As Long, revision As Long, author As Long
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = RouteForAction(db, runID, "ForwardRoute")
    If r!RunState <> "Pending" Then Err.Raise vbObjectError + 4309, "modRouteTemplates.ForwardRoute", "Every step has forwarded. The final reviewer can record local approval or send the report back."
    k = r!CurrentStep: n = RouteStepCount(runID): revision = CurrentRevision(r!ReportID): author = MajorAuthor()
    If RouteRunHolder(r) <> author Then Err.Raise vbObjectError + 4311, "modRouteTemplates.ForwardRoute", "Switch to the contributor holding this report to forward it."
    'Route rows change through recordsets only (see RouteReturnTo).
    Set s = db.OpenRecordset("SELECT * FROM RouteRunStep WHERE RunID=" & runID & " AND StepOrdinal In (" & k & "," & (k + 1) & ")", dbOpenDynaset)
    Do While Not s.EOF
        s.Edit
        If s!StepOrdinal = k Then
            s!StepState = "Accepted": s!DecisionAt = Now(): s!OutgoingRevisionID = revision
            If Len(Trim$(note)) > 0 Then s!DecisionNote = note Else s!DecisionNote = Null
        Else
            s!IncomingRevisionID = revision
        End If
        s.Update: s.MoveNext
    Loop
    s.Close: Set s = Nothing
    r.Edit
    If k >= n Then
        r!RunState = "Accepted": r!ApprovedBy = author: r!ApprovedAt = Now()
    Else
        r!CurrentStep = k + 1
    End If
    r.Update
    RouteEventAdd db, runID, "Forward", k, IIf(k >= n, k, k + 1), revision, note
    r.Close: w.CommitTrans: Exit Sub
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRouteTemplates", "ForwardRoute", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    If Not s Is Nothing Then SwiftClose s, swiftPrimary, "s"
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Sub
'The holder sends the draft back to an earlier step (0 is the requester).
Public Sub SendBackRoute(ByVal runID As Long, ByVal toStep As Long, ByVal note As String)
    Dim db As DAO.Database, w As DAO.Workspace, r As DAO.Recordset
    If Len(Trim$(note)) = 0 Then Err.Raise vbObjectError + 4307, "modRouteTemplates.SendBackRoute", "Say why the report is going back."
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = RouteForAction(db, runID, "SendBackRoute")
    If RouteRunHolder(r) <> MajorAuthor() Then Err.Raise vbObjectError + 4311, "modRouteTemplates.SendBackRoute", "Only the contributor holding this report can send it back."
    If r!RunState = "Pending" And r!CurrentStep = 0 Then Err.Raise vbObjectError + 4319, "modRouteTemplates.SendBackRoute", "The report is already with the requester."
    If toStep < 0 Or toStep >= r!CurrentStep And r!RunState = "Pending" Or toStep > r!CurrentStep Then Err.Raise vbObjectError + 4319, "modRouteTemplates.SendBackRoute", "Choose an earlier step to send the report back to."
    RouteReturnTo db, r, toStep, "Send back", note
    r.Close: w.CommitTrans: Exit Sub
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRouteTemplates", "SendBackRoute", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Sub
'The step a contributor can pull the draft back to: their highest step that has
'already forwarded, or 0 for the requester; -1 when they cannot.
Public Function RoutePullBackStep(ByVal runID As Long, ByVal pid As Long) As Long
    Dim r As DAO.Recordset, limit As Long
    RoutePullBackStep = -1
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RouteRun WHERE RunID=" & runID, dbOpenSnapshot)
    If r.EOF Then r.Close: Exit Function
    If (r!RunState = "Pending" Or r!RunState = "Accepted") And RouteRunHolder(r) <> pid Then
        If r!RunState = "Accepted" Then limit = RouteStepCount(runID) Else limit = r!CurrentStep
        RoutePullBackStep = Nz(MajorScalar("SELECT Max(StepOrdinal) FROM RouteRunStep WHERE RunID=" & runID & " AND AssigneeID=" & pid & " AND StepOrdinal<" & limit), -1)
        If RoutePullBackStep < 0 And r!RequestedBy = pid And limit > 0 Then RoutePullBackStep = 0
    End If
    r.Close
End Function
'The requester or an earlier reviewer takes the draft back; the chain then runs
'upward again from their step.
Public Sub PullBackRoute(ByVal runID As Long, ByVal note As String)
    Dim db As DAO.Database, w As DAO.Workspace, r As DAO.Recordset, target As Long
    If Len(Trim$(note)) = 0 Then Err.Raise vbObjectError + 4307, "modRouteTemplates.PullBackRoute", "Say why you are pulling the report back."
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans
    On Error GoTo Failed
    Set r = RouteForAction(db, runID, "PullBackRoute")
    target = RoutePullBackStep(runID, MajorAuthor())
    If target < 0 Then Err.Raise vbObjectError + 4311, "modRouteTemplates.PullBackRoute", "Only the requester or a reviewer at an earlier step can pull this report back."
    RouteReturnTo db, r, target, "Pull back", note
    r.Close: w.CommitTrans: Exit Sub
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modRouteTemplates", "PullBackRoute", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
    SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Sub
'Who holds a report under review (0 when nobody does), with a description such as
'"with Chief Smith (step 2 of 3)".
Public Function ReviewHolderText(ByVal reportID As Long) As String
    Call ReviewHolder(reportID, ReviewHolderText)
End Function
Public Function ReviewHolder(ByVal reportID As Long, Optional ByRef where As String) As Long
    Dim r As DAO.Recordset
    where = ""
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RouteRun WHERE ReportID=" & reportID & " AND RunState In ('Pending','Accepted') ORDER BY RunID DESC", dbOpenSnapshot)
    If Not r.EOF Then
        ReviewHolder = RouteRunHolder(r)
        where = "with " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & ReviewHolder), "Unknown contributor")
        If r!RunState = "Accepted" Then
            where = where & " for local approval"
        ElseIf r!CurrentStep = 0 Then
            where = where & " (sent back to the requester)"
        Else
            where = where & " (step " & r!CurrentStep & " of " & RouteStepCount(r!RunID) & ")"
        End If
        r.Close: Exit Function
    End If
    r.Close
    Set r = CurrentDb.OpenRecordset("SELECT AssignedTo FROM ReviewTask WHERE ReportID=" & reportID & " AND TaskState='Pending' ORDER BY TaskID DESC", dbOpenSnapshot)
    If Not r.EOF Then ReviewHolder = r!AssignedTo: where = "with " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & ReviewHolder), "Unknown contributor") & " for review"
    r.Close
End Function
'Members-grid saves sit outside the review chain: they are always allowed and are
'logged on the report instead.
Public Sub RouteGridContext(ByVal active As Boolean)
    gridContext = active
End Sub
'Refuses a save by anyone but the contributor holding a report under review.
Public Sub RouteSaveGuard(ByVal reportID As Long)
    Dim holder As Long, where As String
    If gridContext Then Exit Sub
    holder = ReviewHolder(reportID, where)
    If holder = 0 Or holder = MajorAuthor() Then Exit Sub
    'A reviewer later in the chain has more authority and may change the report while
    'it is with someone below them; the change is flagged to the holder.
    If RouteHigherStep(reportID, MajorAuthor()) > 0 Then Exit Sub
    Err.Raise vbObjectError + 4320, "modRouteTemplates.RouteSaveGuard", "This report is " & where & ". Only they, or a reviewer later in the chain, can save changes. Pull it back to edit it."
End Sub
'The contributor's highest step above the current holder's on the report's active
'route, or 0 when they have none.
Public Function RouteHigherStep(ByVal reportID As Long, ByVal pid As Long) As Long
    Dim r As DAO.Recordset, holderStep As Long
    Set r = CurrentDb.OpenRecordset("SELECT * FROM RouteRun WHERE ReportID=" & reportID & " AND RunState='Pending' ORDER BY RunID DESC", dbOpenSnapshot)
    If Not r.EOF Then
        holderStep = r!CurrentStep
        RouteHigherStep = Nz(MajorScalar("SELECT Max(StepOrdinal) FROM RouteRunStep WHERE RunID=" & r!RunID & " AND AssigneeID=" & pid & " AND StepOrdinal>" & holderStep), 0)
    End If
    r.Close
End Function
'The note the last save left for the editor (an override of a later reviewer), read once.
Public Function RouteSaveNotice() As String
    RouteSaveNotice = saveNotice: saveNotice = ""
End Function
'Called inside the save transaction with the fields it changed ("|"-separated).
'Logs Members-grid changes on a report under review, and flags a lower step saving
'over a field a later reviewer last changed in this route.
Public Sub RouteRecordSave(ByVal db As DAO.Database, ByVal reportID As Long, ByVal revisionID As Long, ByVal changedKeys As String)
    Dim r As DAO.Recordset, c As DAO.Recordset, key As Variant, holderStep As Long, kind As String, labels As String, overridden As Object, step As Variant, text As String
    saveNotice = ""
    If Len(changedKeys) = 0 Then Exit Sub
    'Outside review, Members-grid marks replace the working state like any save.
    'Under a route or a pending assignment the change is kept as a visible
    'checkpoint, so the holder and later reviewers can compare it in History.
    If gridContext Then
        If PersistenceCount("RouteRun", "ReportID=" & reportID & " AND RunState In ('Pending','Accepted')") > 0 Or PersistenceCount("ReviewTask", "ReportID=" & reportID & " AND TaskState='Pending'") > 0 Then
            db.Execute "UPDATE RevisionRetention SET RetentionKind='Checkpoint',CheckpointReason='Members grid change' WHERE RevisionID=" & revisionID & " AND RetentionKind In ('Working','Protected')", dbFailOnError
        End If
    End If
    kind = Nz(MajorScalar("SELECT FormKind FROM Report WHERE ReportID=" & reportID), "")
    For Each key In Split(changedKeys, "|")
        If Len(labels) > 0 Then labels = labels & ", "
        labels = labels & RouteFieldLabel(CStr(key), kind)
    Next
    Set r = db.OpenRecordset("SELECT * FROM RouteRun WHERE ReportID=" & reportID & " AND RunState In ('Pending','Accepted') ORDER BY RunID DESC", dbOpenSnapshot)
    If r.EOF Then
        r.Close
        If gridContext Then
            Set r = db.OpenRecordset("SELECT TaskID FROM ReviewTask WHERE ReportID=" & reportID & " AND TaskState='Pending'", dbOpenSnapshot)
            Do While Not r.EOF
                db.Execute "INSERT INTO ReviewTaskEvent (TaskID,AuthorProfileID,EventAt,EventKind,Notes) VALUES (" & r!TaskID & "," & MajorAuthor() & ",Now()," & QuoteSQL("Grid change") & "," & QuoteSQL(Left$("Changed from the Members grid: " & labels, 4000)) & ")", dbFailOnError
                r.MoveNext
            Loop
            r.Close
        End If
        Exit Sub
    End If
    'The holder's own changes, from the editor or the grid, need no flag.
    If gridContext And RouteRunHolder(r) <> MajorAuthor() Then
        RouteEventAdd db, r!RunID, "Grid change", Null, Null, revisionID, "Changed from the Members grid: " & labels
        r.Close: Exit Sub
    End If
    If r!RunState = "Accepted" Then holderStep = RouteStepCount(r!RunID) Else holderStep = r!CurrentStep
    If RouteRunHolder(r) <> MajorAuthor() And Not gridContext Then
        step = MajorScalar("SELECT Max(StepOrdinal) FROM RouteRunStep WHERE RunID=" & r!RunID & " AND AssigneeID=" & MajorAuthor() & " AND StepOrdinal>" & holderStep)
        RouteEventAdd db, r!RunID, "Higher change", step, holderStep, revisionID, "Changed by " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & MajorAuthor()), "Unknown contributor") & " (step " & Nz(step, "?") & "): " & labels
        r.Close: Exit Sub
    End If
    If gridContext Then r.Close: Exit Sub
    Set overridden = CreateObject("Scripting.Dictionary")
    For Each key In Split(changedKeys, "|")
        'The field's previous change in this route, and whether a later step made it.
        Set c = db.OpenRecordset("SELECT TOP 1 v.AuthorProfileID FROM RevisionChange AS g INNER JOIN ReportRevision AS v ON g.RevisionID=v.RevisionID WHERE v.ReportID=" & reportID & " AND g.FieldKey=" & QuoteSQL(CStr(key)) & " AND v.RevisionID<>" & revisionID & " AND v.CommittedAt>=" & SQLStamp(r!RequestedAt) & " ORDER BY v.RevisionID DESC", dbOpenSnapshot)
        If Not c.EOF Then
            If c!AuthorProfileID <> MajorAuthor() Then
                step = MajorScalar("SELECT Max(StepOrdinal) FROM RouteRunStep WHERE RunID=" & r!RunID & " AND AssigneeID=" & c!AuthorProfileID & " AND StepOrdinal>" & holderStep)
                If Not IsNull(step) Then
                    text = Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & c!AuthorProfileID), "Unknown contributor") & " (step " & step & ")"
                    If overridden.Exists(text) Then overridden(text) = overridden(text) & ", " & RouteFieldLabel(CStr(key), kind) Else overridden.Add text, RouteFieldLabel(CStr(key), kind)
                End If
            End If
        End If
        c.Close
    Next
    For Each step In overridden.Keys
        text = "Overrides a change by " & step & ": " & overridden(step)
        RouteEventAdd db, r!RunID, "Override", holderStep, Null, revisionID, text
        If Len(saveNotice) > 0 Then saveNotice = saveNotice & " "
        saveNotice = saveNotice & text & "."
    Next
    r.Close
End Sub
Private Function RouteFieldLabel(ByVal key As String, ByVal kind As String) As String
    On Error Resume Next
    RouteFieldLabel = UXFieldCaption(key, kind)
    If Len(RouteFieldLabel) = 0 Then RouteFieldLabel = key
End Function
Private Function SQLStamp(ByVal value As Variant) As String
    If IsNull(value) Then SQLStamp = "#1900-01-01#" Else SQLStamp = "#" & Format$(value, "yyyy-mm-dd hh:nn:ss") & "#"
End Function
