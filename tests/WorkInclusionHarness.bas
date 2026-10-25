Attribute VB_Name = "WorkInclusionHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8792, , message
 checks = checks + 1
End Sub
Private Function ReadRows() As String
 Dim r As DAO.Recordset, f As DAO.Field
 Set r = CurrentDb.OpenRecordset(UXWorkSQL(MajorAuthor(), False, True), dbOpenSnapshot)
 Do While Not r.EOF
  For Each f In r.Fields
   If IsNull(f.Value) Then ReadRows = ReadRows & "NULL|" Else ReadRows = ReadRows & Len(CStr(f.Value)) & ":" & CStr(f.Value) & "|"
  Next
  ReadRows = ReadRows & vbLf: r.MoveNext
 Loop
 r.Close
End Function
Private Function QueueSample(ByRef value As String) As Long
 Dim started As Double
 started = Timer
 UXPrepareWorkRows MajorAuthor(), False
 value = ReadRows(): QueueSample = CLng((Timer - started) * 1000)
End Function
Private Function CountSample(ByRef value As Long) As Long
 Dim started As Double
 started = Timer
 value = UXWorkCount()
 CountSample = CLng((Timer - started) * 1000)
End Function
Public Function TestWorkInclusion(ByVal output As String) As String
 Dim target As Variant, i As Long, id As Long, p As Long, g As Long, template As Long, ignored As Long, trial As Long, before As String, after As String, firstCount As Long, secondCount As Long, firstMs As Long, secondMs As Long, file As Integer, revisions As Long, snapshots As Long, visible As Object, failure As Long, message As String
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC WORK INCLUSION", "Reviewer"
 p = MajorAuthor(): g = CreateSummaryGroup("Work scale", "EVAL", "E5")
 ignored = AddContributor("Unrelated reviewer")
 template = CreateRouteTemplate("Scale route", CStr(p), "Read saved report")
 file = FreeFile: Open output For Output As #file
 Print #file, "Reports,Area,Trial,CurrentMs,RepeatMs"
 For Each target In Array(25, 100)
  For i = DCount("*", "Report") + 1 To CLng(target)
   id = CommitDraft(0, "EVAL", "Synthetic " & i, "Member " & Format$(i, "000"), #1/1/2026#, #9/30/2026#, "")
   AddGroupMember g, id, "E5"
   ignored = AssignReview(id, CurrentRevision(id), p, Null, "Review")
   ignored = StartRoute(id, template, g)
   ignored = AddReviewComment(id, CurrentRevision(id), "Synthetic comment")
  Next
  revisions = DCount("*", "ReportRevision"): snapshots = DCount("*", "GroupSnapshot") + DCount("*", "ComparisonSet")
  For trial = 1 To 3
   firstMs = QueueSample(before): secondMs = QueueSample(after)
   Call CountSample(firstCount): Call CountSample(secondCount)
   Check StrComp(before, after, vbBinaryCompare) = 0, "Queue contents/order differ"
   Set visible = UXWorkVisibleRows(MajorAuthor())
   Check firstCount = secondCount And secondCount = visible.Count, "Count differs"
   Print #file, target & ",Queue," & trial & "," & firstMs & "," & secondMs
  Next
  Check DCount("*", "ReportRevision") = revisions, "Reading wrote revisions"
  Check DCount("*", "GroupSnapshot") + DCount("*", "ComparisonSet") = snapshots, "Reading wrote snapshots"
 Next
 Close #file
 CheckCaseAndMissingRevision
 TestWorkInclusion = "PASS: " & checks & " scale equivalence/count/read-only assertions": Exit Function
Failed:
 failure = Err.Number: message = Err.Description
 On Error Resume Next: Close #file
 TestWorkInclusion = "FAIL: " & failure & " " & message
End Function
Public Function TestNoVisibilityCallbacks() As String
 Dim n As Long, r As DAO.Recordset, visible As Object
 On Error GoTo Failed
 Set visible = UXWorkVisibleRows(MajorAuthor()): n = UXWorkCount()
 Check n = visible.Count, "Callback-free count differs"
 UXPrepareWorkRows MajorAuthor(), False
 Set r = CurrentDb.OpenRecordset(UXWorkSQL(MajorAuthor(), False, True), dbOpenSnapshot)
 If Not r.EOF Then r.MoveLast
 Check r.RecordCount = n, "Callback-free queue differs"
 r.Close
 TestNoVisibilityCallbacks = "PASS: inclusion, count and prepared queue do not call single-item UXWorkVisible": Exit Function
Failed:
 TestNoVisibilityCallbacks = "FAIL: " & Err.Number & " " & Err.Description
End Function

Private Function VisibilityError(ByVal batch As Boolean, ByVal id As Long) As Long
 Dim visible As Object, ignored As Boolean
 On Error GoTo Failed
 If batch Then Set visible = UXWorkVisibleRows(MajorAuthor()) Else ignored = UXWorkVisible("Assignment", id, MajorAuthor())
 Exit Function
Failed:
 VisibilityError = Err.Number
End Function
Private Function RetiredResolveError(ByVal thread As Long) As Long
 Dim ignored As Boolean
 On Error GoTo Failed
 ignored = SetReviewThreadState(thread, True)
 Exit Function
Failed:
 RetiredResolveError = Err.Number
End Function
' A review thread imported from before report messages, with a lowercase historical event.
Private Function LegacyThread(ByVal report As Long) As Long
 CurrentDb.Execute "INSERT INTO ReviewComment (CommentGuid,ReportID,RevisionID,AuthorProfileID,CommentAt,CommentBody) VALUES (" & QuoteSQL(NewIdentity()) & "," & report & "," & CurrentRevision(report) & "," & MajorAuthor() & ",Now(),'Synthetic legacy thread')", dbFailOnError
 LegacyThread = DMax("CommentID", "ReviewComment")
 CurrentDb.Execute "INSERT INTO ReviewCommentEvent (ThreadID,AuthorProfileID,EventAt,EventKind) VALUES (" & LegacyThread & "," & MajorAuthor() & ",Now(),'resolved')", dbFailOnError
End Function
Private Sub CheckCaseAndMissingRevision()
 Dim task As Long, run As Long, report As Long, thread As Long, legacy As Long, reviewer As Long, approver As Long, revision As Long, pid As Variant, kind As Variant, id As Long, visible As Object
 task = DMin("TaskID", "ReviewTask"): report = DLookup("ReportID", "ReviewTask", "TaskID=" & task)
 ' Review comments are report messages; My Work addresses a message thread by its negated MessageID.
 run = DMin("RunID", "RouteRun", "ReportID=" & report): thread = -Nz(DMin("MessageID", "ReportMessage", "ReportID=" & report & " AND ParentMessageID Is Null"), 0)
 Check thread < 0, "Review comment did not start a report message thread"
 legacy = LegacyThread(report)
 reviewer = DMax("UserProfileID", "UserProfile"): approver = AddContributor("Case-sensitive route fixture")
 CurrentDb.Execute "UPDATE ReviewTask SET TaskState='pending',AssignedTo=" & reviewer & " WHERE TaskID=" & task, dbFailOnError
 CurrentDb.Execute "UPDATE RouteRun SET RunState='accepted',ApprovedBy=" & approver & " WHERE RunID=" & run, dbFailOnError
 Check RetiredResolveError(thread) = vbObjectError + 1706, "Retired thread resolve/reopen accepted"
 For Each pid In Array(MajorAuthor(), reviewer, approver, 0)
  Set visible = UXWorkVisibleRows(CLng(pid))
  For Each kind In Array("Assignment", "Route", "Comment", "LegacyComment")
   Select Case kind
    Case "Assignment": id = task
    Case "Route": id = run
    Case "Comment": id = thread
    Case "LegacyComment": id = legacy
   End Select
   Check visible.Exists(Replace(kind, "Legacy", "") & ":" & id) = UXWorkVisible(Replace(kind, "Legacy", ""), id, CLng(pid)), "Mixed-case inclusion differs: " & kind
  Next
 Next
 Check UXThreadOpen(legacy), "Event comparison stopped being binary"
 ' No saved revision: neither the authoritative head nor the derived cache pointer exists.
 ' The single-item oracle reads ReportHead (8900); batched rows read the derived cache (4002).
 revision = CurrentRevision(report)
 CurrentDb.Execute "DELETE FROM ReportHead WHERE ReportID=" & report, dbFailOnError
 CurrentDb.Execute "UPDATE Report SET CurrentRevisionID=Null WHERE ReportID=" & report, dbFailOnError
 Check VisibilityError(False, task) = vbObjectError + 8900, "Legacy missing revision fixture invalid"
 Check VisibilityError(True, task) = vbObjectError + 4002, "Batched missing revision error changed"
 CurrentDb.Execute "INSERT INTO ReportHead (ReportID,RevisionID) VALUES (" & report & "," & revision & ")", dbFailOnError
 CurrentDb.Execute "UPDATE Report SET CurrentRevisionID=" & revision & " WHERE ReportID=" & report, dbFailOnError
 Check CurrentRevision(report) = revision, "Missing revision fixture not restored"
End Sub
Public Function TestVisibilityStub() As Boolean
 Dim ignored As Boolean
 On Error GoTo Blocked
 ignored = UXWorkVisible("Assignment", 1, MajorAuthor())
 Exit Function
Blocked:
 TestVisibilityStub = (Err.Number = vbObjectError + 8793)
End Function
