Attribute VB_Name = "GridReadHarness"
Option Explicit
Private checks As Long
Private loadMessage As String
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8671, , message
 checks = checks + 1
End Sub
Private Function LoadError() As Long
 On Error GoTo Failed
 UXGridLoad
 Exit Function
Failed:
 LoadError = Err.Number: loadMessage = Err.Description
End Function
Private Function Capture() As Collection
 Dim result As New Collection, rows As DAO.Recordset, field As DAO.Field
 Set rows = Forms!frmUXComparison!uxTraitGrid.Form.RecordsetClone
 If Not rows.EOF Then rows.MoveFirst
 Do While Not rows.EOF
  For Each field In rows.Fields
   If IsNull(field.Value) Then result.Add "NULL" Else result.Add "VALUE:" & CStr(field.Value)
  Next
  rows.MoveNext
 Loop
 rows.Close: Set Capture = result
End Function
Private Sub Compare(ByVal label As String, ByVal file As Integer)
 Dim expected As Collection, actual As Collection, i As Long, started As Double, oldMs As Long, newMs As Long
 Dim revisions As Long, snapshots As Long
 revisions = DCount("*", "ReportRevision"): snapshots = DCount("*", "ComparisonSet")
 started = Timer: Check UXGridLoadLegacy(), "Legacy load": oldMs = CLng((Timer - started) * 1000)
 Set expected = Capture()
 started = Timer: Check UXGridLoad(), "New load": newMs = CLng((Timer - started) * 1000)
 Set actual = Capture()
 Check expected.Count = actual.Count, label & " field count differs"
 For i = 1 To expected.Count: Check expected(i) = actual(i), label & " field differs at " & i: Next
 Check revisions = DCount("*", "ReportRevision") And snapshots = DCount("*", "ComparisonSet"), "Reads wrote history"
 Print #file, label & "," & oldMs & "," & newMs & "," & actual.Count
End Sub
Public Function TestGridReads(ByVal output As String) As String
 Dim id As Long, g As Long, i As Long, m As Long, marks As Object, b As Long, prior As Long, file As Integer, meeting As Long, failure As Long, message As String
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC GRID READS", "Tester"
 Set marks = CreateObject("Scripting.Dictionary")
 For b = 33 To 39: marks.Add "Trait" & b, "4.0": Next
 g = CreateSummaryGroup("Batch grid", "EVAL", "E5")
 file = FreeFile: Open output For Output As #file: Print #file, "Case,LegacyMs,NewMs,Fields"
 For i = 1 To 100
  id = CommitDraft(0, "EVAL", "Synthetic " & i, "Member " & i, #1/1/2026#, #9/30/2026#, "", "Test", "", , Array("4.0", "4.0", "4.0", "4.0", "4.0", "4.0", "4.0"), Array("Observed", "Promotable"))
  m = CLng(DLookup("MemberID", "MemberReport", "ReportID=" & id))
  LinkReportDirectoryMember id, m, CurrentRevision(id), CLng(DLookup("VersionCounter", "CommandMember", "MemberID=" & m)), "Synthetic confirmation"
  prior = SavePriorEvaluation(0, 0, m, "EVAL", "IT2", #1/1/2025#, #12/31/2025#, "Promotable", "Observed", marks, "Synthetic prior")
  AddGroupMember g, id, "E5"
  If i = 25 Or i = 50 Or i = 100 Then
   DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = g: Forms!frmMajorWorkbench!uxMode = "Members"
   Check UXOpenComparison(), "Open live"
   Compare "Live" & i, file
   Check UXCloseComparison(), "Close live"
   Forms!frmMajorWorkbench!uxMode = "Comparison": Check UXOpenComparison(), "Open historical"
   meeting = Forms!frmUXComparison!txtMeeting
   ProposeMeetingField meeting, id, "Trait33", "3.0"
   ProposeMeetingField meeting, id, "AssignmentRecommendationA", "STAFF"
   Compare "Historical" & i, file
   If i = 100 Then
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & CurrentRevision(id) & " AND FieldKey='Trait39'", dbFailOnError
    failure = LoadError(): Check failure = vbObjectError + 4007, "Missing revision field error changed: " & failure & " " & loadMessage
    Check Forms!frmUXComparison!uxTraitGrid.Form.Recordset Is Nothing, "Failed cache remained bound"
    Check UXGridPendingCount() = 0, "Failed load created pending edits"
    CurrentDb.Execute "INSERT INTO RevisionValue (RevisionID,FieldKey,ValueText,IsNullValue) VALUES (" & CurrentRevision(id) & ",'Trait39','4.0',False)", dbFailOnError
    Check UXGridLoad(), "Reload after failed cache failed"
   End If
   Check UXCloseComparison(), "Close historical"
   DoCmd.Close acForm, "frmMajorWorkbench", acSaveNo
  End If
 Next
 Close #file
 TestGridReads = "PASS: " & checks & " exact grid field, order, live/historical proposal and read-only assertions": Exit Function
Failed:
 failure = Err.Number: message = Err.Description
 On Error Resume Next: Close #file
 TestGridReads = "FAIL after " & checks & ": " & failure & " " & message
End Function
