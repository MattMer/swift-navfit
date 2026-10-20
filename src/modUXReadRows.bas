Attribute VB_Name = "modUXReadRows"
Option Compare Binary
Option Explicit
' Display snapshots only. Every refresh replaces the complete map; decision services
' continue to use live authoritative reads. No workspace/cache tables are written.
Private workRows As Object
Private reportRows As Object
Private Function ReadMap() As Object
    Set ReadMap = CreateObject("Scripting.Dictionary")
End Function
Private Function ReadRouteCurrent(ByVal r As DAO.Recordset, ByVal stamps As Object) As Boolean
    Dim key As String
    ReadRouteCurrent = (r!RevisionID = Nz(r!SavedRevision, 0)) And Not Nz(r!ReportArchived, True)
    If Not IsNull(r!GroupID) Then
        key = CStr(r!GroupID)
        If Not stamps.Exists(key) Then stamps.Add key, GroupStamp(r!GroupID, False)
        If StrComp(Nz(r!GroupStamp, ""), stamps(key), vbBinaryCompare) <> 0 Then ReadRouteCurrent = False
    End If
End Function

' Fresh inclusion keys only: no formatted rows and no persistent cache. The public
' single-item visibility service remains available for authoritative UI decisions.
Public Function UXWorkVisibleRows(ByVal pid As Long) As Object
    Dim db As DAO.Database, r As DAO.Recordset, rows As Object, steps As Object, threads As Object, participants As Object, authors As Object, stamps As Object
    Dim sql As String, token As String, id As Long, assignee As Long, visible As Boolean, current As Boolean
    Set db = CurrentDb: Set rows = ReadMap(): Set steps = ReadMap(): Set threads = ReadMap()
    Set participants = ReadMap(): Set authors = ReadMap(): Set stamps = ReadMap()
    sql = "SELECT ReportID FROM ReportRevision WHERE AuthorProfileID=" & pid
    sql = sql & " UNION SELECT ReportID FROM ReviewTask WHERE TaskState='Pending' AND (AssignedTo=" & pid & " OR RequestedBy=" & pid & ")"
    sql = sql & " UNION SELECT ReportID FROM RouteRun WHERE RunState In ('Pending','Accepted') AND (RequestedBy=" & pid & " OR ApprovedBy=" & pid & ")"
    sql = sql & " UNION SELECT u.ReportID FROM RouteRunStep AS s INNER JOIN RouteRun AS u ON s.RunID=u.RunID WHERE u.RunState='Pending' AND s.StepOrdinal=u.CurrentStep AND s.AssigneeID=" & pid
    Set r = db.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not r.EOF: participants(CStr(r!ReportID)) = True: r.MoveNext: Loop
    r.Close
    Set r = db.OpenRecordset("SELECT CommentID,ThreadID FROM ReviewComment WHERE AuthorProfileID=" & pid, dbOpenSnapshot)
    Do While Not r.EOF
        If IsNull(r!ThreadID) Then token = CStr(r!CommentID) Else token = CStr(r!ThreadID)
        authors(token) = True: r.MoveNext
    Loop
    r.Close
    Set r = db.OpenRecordset("SELECT ThreadID,EventKind FROM ReviewCommentEvent ORDER BY CommentEventID", dbOpenSnapshot)
    Do While Not r.EOF: threads(CStr(r!ThreadID)) = (r!EventKind <> "Resolved"): r.MoveNext: Loop
    r.Close
    Set r = db.OpenRecordset("SELECT RunID,StepOrdinal,AssigneeID FROM RouteRunStep", dbOpenSnapshot)
    Do While Not r.EOF: steps(r!RunID & ":" & r!StepOrdinal) = Nz(r!AssigneeID, 0): r.MoveNext: Loop
    r.Close
    Set r = db.OpenRecordset("SELECT t.TaskID,t.RevisionID,t.AssignedTo,t.RequestedBy,t.TaskState,p.CurrentRevisionID AS SavedRevision FROM ReviewTask AS t LEFT JOIN Report AS p ON t.ReportID=p.ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        ' VBA's original And/Or evaluates CurrentRevision even for completed tasks.
        If Nz(r!SavedRevision, 0) = 0 Then Err.Raise vbObjectError + 4002, , "Select a saved report."
        If r!TaskState = "Pending" And (r!AssignedTo = pid Or (r!RequestedBy = pid And r!RevisionID <> r!SavedRevision)) Then rows("Assignment:" & r!TaskID) = True
        r.MoveNext
    Loop
    r.Close
    Set r = db.OpenRecordset("SELECT u.*,p.CurrentRevisionID AS SavedRevision,p.IsArchived AS ReportArchived FROM RouteRun AS u LEFT JOIN Report AS p ON u.ReportID=p.ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        id = r!RunID: visible = False
        If r!RunState = "Pending" Or r!RunState = "Accepted" Then
            If Nz(r!SavedRevision, 0) = 0 Then Err.Raise vbObjectError + 4002, , "Select a saved report."
            current = ReadRouteCurrent(r, stamps)
            If r!RunState = "Pending" Then
                assignee = 0
                If steps.Exists(id & ":" & r!CurrentStep) Then assignee = steps(id & ":" & r!CurrentStep)
                visible = assignee = pid Or (r!RequestedBy = pid And Not current)
            Else
                visible = Nz(r!ApprovedBy, 0) = pid Or (r!RequestedBy = pid And Not current)
            End If
        End If
        If visible Then rows("Route:" & id) = True
        r.MoveNext
    Loop
    r.Close
    Set UXWorkVisibleRows = rows
End Function
Public Sub UXPrepareWorkRows(ByVal pid As Long, ByVal history As Boolean)
    Dim rows As Object, profiles As Object, steps As Object, counts As Object, threads As Object, stamps As Object, visibleRows As Object
    Dim db As DAO.Database, r As DAO.Recordset, kind As Variant, table As String, key As String, sql As String
    Dim id As Long, owner As Long, label As String, status As String, action As String, due As String, member As String, stale As Boolean, openThread As Boolean, visible As Boolean
    Set rows = ReadMap(): Set profiles = ReadMap(): Set steps = ReadMap(): Set counts = ReadMap(): Set threads = ReadMap(): Set stamps = ReadMap()
    If Not history Then Set visibleRows = UXWorkVisibleRows(pid)
    Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT UserProfileID,DisplayName FROM UserProfile", dbOpenSnapshot)
    Do While Not r.EOF
        profiles(CStr(r!UserProfileID)) = Nz(r!DisplayName, "Unknown contributor"): r.MoveNext
    Loop
    r.Close
    Set r = db.OpenRecordset("SELECT RunID,StepOrdinal,AssigneeID FROM RouteRunStep", dbOpenSnapshot)
    Do While Not r.EOF
        key = CStr(r!RunID): steps(key & ":" & r!StepOrdinal) = Nz(r!AssigneeID, 0)
        If Not counts.Exists(key) Then counts(key) = 0
        counts(key) = counts(key) + 1: r.MoveNext
    Loop
    r.Close
    Set r = db.OpenRecordset("SELECT ThreadID,EventKind FROM ReviewCommentEvent ORDER BY CommentEventID", dbOpenSnapshot)
    Do While Not r.EOF
        threads(CStr(r!ThreadID)) = (r!EventKind <> "Resolved"): r.MoveNext
    Loop
    r.Close
    For Each kind In Array("Assignment", "Route")
        Select Case kind
            Case "Assignment": table = "ReviewTask": key = "TaskID"
            Case "Route": table = "RouteRun": key = "RunID"
            Case "Comment": table = "ReviewComment": key = "CommentID"
        End Select
        sql = "SELECT w.*,p.OfficialName,p.MemberName,p.FormKind,p.PeriodStart,p.PeriodEnd,p.CurrentRevisionID AS SavedRevision,p.IsArchived AS ReportArchived,p.WorkflowState AS ReportWorkflow FROM " & table & " AS w LEFT JOIN Report AS p ON w.ReportID=p.ReportID"
        If kind = "Comment" Then sql = sql & " WHERE w.ThreadID Is Null"
        Set r = db.OpenRecordset(sql, dbOpenSnapshot)
        Do While Not r.EOF
            id = r.Fields(key): owner = 0: label = "": status = "": action = "": due = "Not set"
            visible = history
            If Not history Then visible = visibleRows.Exists(kind & ":" & id)
            If visible Then
                member = UXIdentity(r!OfficialName, r!MemberName) & " | " & r!FormKind & " | " & Nz(Format(r!PeriodStart, "yyyy-mm-dd"), "No start date") & " to " & Nz(Format(r!PeriodEnd, "yyyy-mm-dd"), "No end date")
                Select Case kind
                    Case "Assignment"
                        label = "Assigned review": owner = r!AssignedTo: status = r!TaskState
                        If status = "Pending" Then
                            stale = r!RevisionID <> Nz(r!SavedRevision, 0)
                            If stale Then status = "Stale": action = "Requester: cancel and reassign" Else action = "Open report; accept or return"
                            If Not IsNull(r!DueDate) Then
                                If r!DueDate < Date Then status = status & " / overdue"
                            End If
                        Else
                            action = "View review history"
                        End If
                        due = Nz(Format(r!DueDate, "yyyy-mm-dd"), "Not set")
                    Case "Route"
                        If Not counts.Exists(CStr(id)) Then counts(CStr(id)) = 0
                        label = "Route: step " & r!CurrentStep & " of " & counts(CStr(id)): status = r!RunState
                        If steps.Exists(id & ":" & r!CurrentStep) Then owner = steps(id & ":" & r!CurrentStep)
                        If status = "Pending" Or status = "Accepted" Then
                            stale = Not ReadRouteCurrent(r, stamps)
                            If stale Then
                                status = "Stale": action = "Requester: restart route": owner = r!RequestedBy
                            ElseIf status = "Accepted" Then
                                label = "Route: all steps accepted": action = "Final reviewer: local approval": owner = Nz(r!ApprovedBy, 0)
                            Else
                                action = "Assigned reviewer: accept or return"
                            End If
                        Else
                            action = "View route history"
                        End If
                    Case "Comment"
                        label = "Comment thread": owner = r!AuthorProfileID: openThread = True
                        If threads.Exists(CStr(id)) Then openThread = threads(CStr(id))
                        If openThread Then status = "Unresolved": action = "Read, reply or resolve" Else status = "Resolved": action = "View thread history"
                        If r!RevisionID <> Nz(r!SavedRevision, 0) Then status = status & " / earlier revision"
                End Select
                If Nz(r!ReportArchived, True) Then
                    status = status & " / archived": action = "Recover report before review activity"
                ElseIf Nz(r!ReportWorkflow, "") = "Approved" And kind <> "Comment" Then
                    status = status & " / local approval": action = "Reopen local approval before decisions"
                End If
                If Not profiles.Exists(CStr(owner)) Then profiles(CStr(owner)) = "Unknown contributor"
                rows.Add kind & ":" & id, Array(member, label, profiles(CStr(owner)), due, status, action)
            End If
            r.MoveNext
        Loop
        r.Close
    Next
    Set workRows = rows
End Sub
Public Function UXWorkRowExists(ByVal kind As String, ByVal id As Long) As Boolean
    If workRows Is Nothing Then Exit Function
    UXWorkRowExists = workRows.Exists(kind & ":" & id)
End Function
Public Function UXWorkRowCell(ByVal kind As String, ByVal id As Long, ByVal column As Long) As String
    Dim row As Variant
    If Not UXWorkRowExists(kind, id) Then Exit Function
    row = workRows(kind & ":" & id): UXWorkRowCell = row(column)
End Function
Private Sub ReadReportState(ByVal rows As Object, ByVal id As Long, ByVal state As String)
    Dim row As Variant, prior As String
    If Not rows.Exists(CStr(id)) Then Exit Sub
    row = rows(CStr(id)): prior = row(1)
    If prior = "Archived" Or prior = "Local approval" Or prior = "Review needs restart" Then Exit Sub
    If state = "Returned" And prior = "In review" Then Exit Sub
    row(1) = state: rows(CStr(id)) = row
End Sub
Public Sub UXPrepareReportRows()
    Dim rows As Object, traits As Object, values As Object, stamps As Object, db As DAO.Database, r As DAO.Recordset
    Dim row As Variant, id As Variant, key As Variant, value As Variant, total As Double, count As Long, complete As Boolean, fields As String
    Set rows = ReadMap(): Set traits = ReadMap(): Set stamps = ReadMap(): Set db = CurrentDb
    Set r = db.OpenRecordset("SELECT p.*,v.RevisionNumber FROM Report AS p LEFT JOIN ReportRevision AS v ON p.CurrentRevisionID=v.RevisionID", dbOpenSnapshot)
    Do While Not r.EOF
        row = Array(UXIdentity(r!OfficialName, r!MemberName), "Draft", "Unavailable", "", r!RevisionNumber.Value, Nz(r!CurrentRevisionID, 0), r!ObservationStatus.Value)
        If Nz(r!IsArchived, False) Then
            row(1) = "Archived"
        ElseIf Nz(r!WorkflowState, "Draft") = "Approved" Then
            row(1) = "Local approval"
        End If
        rows.Add CStr(r!ReportID), row: r.MoveNext
    Loop
    r.Close
    For Each key In TraitKeys()
        If Len(fields) > 0 Then fields = fields & ","
        fields = fields & QuoteSQL(CStr(key))
    Next
    Set r = db.OpenRecordset("SELECT p.ReportID,v.FieldKey,v.ValueText,v.IsNullValue FROM Report AS p INNER JOIN RevisionValue AS v ON p.CurrentRevisionID=v.RevisionID WHERE v.FieldKey In (" & fields & ")", dbOpenSnapshot)
    Do While Not r.EOF
        id = CStr(r!ReportID)
        If Not traits.Exists(id) Then traits.Add id, ReadMap()
        Set values = traits(id)
        If r!IsNullValue Then values(CStr(r!FieldKey)) = Null Else values(CStr(r!FieldKey)) = r!ValueText.Value
        r.MoveNext
    Loop
    r.Close
    For Each id In rows.Keys
        row = rows(id): total = 0: count = 0: complete = False
        If Nz(row(6), "") = "NOB" Then
            row(2) = "NOB"
        Else
            If row(5) = 0 Then Err.Raise vbObjectError + 4002, , "Select a saved report."
            If traits.Exists(id) Then
                Set values = traits(id): complete = True
                For Each key In TraitKeys()
                    If Not values.Exists(CStr(key)) Then
                        complete = False
                    Else
                        value = values(CStr(key))
                        If Not IsNull(value) Then
                            If value <> "" And value <> "NOB" Then total = total + Val(value): count = count + 1
                        End If
                    End If
                Next
            End If
            If complete And count > 0 Then row(2) = Format$(total / count, "0.00")
        End If
        rows(id) = row
    Next
    Set r = db.OpenRecordset("SELECT m.ReportID,g.Title FROM GroupMember AS m INNER JOIN SummaryGroup AS g ON m.GroupID=g.GroupID ORDER BY g.Title,g.GroupID", dbOpenSnapshot)
    Do While Not r.EOF
        id = CStr(r!ReportID)
        If rows.Exists(id) Then
            row = rows(id)
            If Len(row(3)) > 0 Then row(3) = row(3) & "; "
            row(3) = row(3) & r!Title: rows(id) = row
        End If
        r.MoveNext
    Loop
    r.Close
    For Each id In rows.Keys
        row = rows(id): If Len(row(3)) = 0 Then row(3) = "Ungrouped"
        rows(id) = row
    Next
    Set r = db.OpenRecordset("SELECT t.ReportID,t.RevisionID,t.TaskState,p.CurrentRevisionID FROM ReviewTask AS t INNER JOIN Report AS p ON t.ReportID=p.ReportID WHERE t.TaskState In ('Pending','Returned')", dbOpenSnapshot)
    Do While Not r.EOF
        If r!TaskState = "Pending" Then
            If r!RevisionID <> Nz(r!CurrentRevisionID, 0) Then ReadReportState rows, r!ReportID, "Review needs restart" Else ReadReportState rows, r!ReportID, "In review"
        ElseIf r!RevisionID = Nz(r!CurrentRevisionID, 0) Then
            ReadReportState rows, r!ReportID, "Returned"
        End If
        r.MoveNext
    Loop
    r.Close
    Set r = db.OpenRecordset("SELECT u.*,p.CurrentRevisionID AS SavedRevision,p.IsArchived AS ReportArchived FROM RouteRun AS u LEFT JOIN Report AS p ON u.ReportID=p.ReportID WHERE u.RunState In ('Pending','Accepted')", dbOpenSnapshot)
    Do While Not r.EOF
        If Not ReadRouteCurrent(r, stamps) Then
            ReadReportState rows, r!ReportID, "Review needs restart"
        ElseIf r!RunState = "Pending" Then
            ReadReportState rows, r!ReportID, "In review"
        End If
        r.MoveNext
    Loop
    r.Close
    Set reportRows = rows
End Sub
Public Function UXReportRowCell(ByVal id As Long, ByVal column As Long) As Variant
    Dim row As Variant
    If reportRows Is Nothing Then Err.Raise 5, , "Refresh Reports before reading display values."
    If Not reportRows.Exists(CStr(id)) Then Err.Raise 5, , "Report display changed; refresh Reports."
    row = reportRows(CStr(id)): UXReportRowCell = row(column)
End Function
