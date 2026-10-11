Attribute VB_Name = "TraitQueueHarness"
Option Explicit
Public Function TestTraitQueue() As String
 Dim groupID As Long, first As Long, second As Long, revision As Long, rows As Access.Form, r As DAO.Recordset
 On Error GoTo Failed
 ConfigureWorkspace "Queue tests", "Tester"
 groupID = CreateSummaryGroup("Queued E6", "EVAL", "E6")
 first = CreateGroupDraft(groupID, GroupStamp(groupID), UXDefaultsStamp(groupID), 0, 0, "Synthetic first", "FIRST, SYNTHETIC", "IT1", "", Null, Null)
 second = CreateGroupDraft(groupID, GroupStamp(groupID), UXDefaultsStamp(groupID), 0, 0, "Synthetic second", "SECOND, SYNTHETIC", "IT1", "", Null, Null)
 DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = groupID
 UXWorkbenchMode "Members": UXOpenComparison
 Set rows = Forms!frmUXComparison!uxTraitGrid.Form
 Set r = rows.RecordsetClone: r.FindFirst "ReportID=" & first: rows.Bookmark = r.Bookmark: r.Close
 revision = CurrentRevision(first)
 rows!Trait37 = "1.0": UXGridQueue "Trait37"
 rows!Trait37 = "2.0": UXGridQueue "Trait37"
 If rows.Dirty Then rows.Dirty = False
 Set r = rows.RecordsetClone: r.FindFirst "ReportID=" & second: rows.Bookmark = r.Bookmark: r.Close
 rows!Trait38 = "1.0": UXGridQueue "Trait38"
 If Not UXGridDrain() Then Err.Raise vbObjectError + 7920, , Forms!frmUXComparison!lblStatus.Caption
 If DLookup("Trait37", "Report", "ReportID=" & first) <> "2.0" Or DLookup("Trait38", "Report", "ReportID=" & second) <> "1.0" Then Err.Raise vbObjectError + 7921, , "Queued edit saved against wrong row"
 If DCount("*", "ReportRevision", "ReportID=" & first & " AND RevisionID>" & revision) <> 1 Then Err.Raise vbObjectError + 7922, , "Repeated same cell created excess revisions"
 Set r = rows.RecordsetClone: r.FindFirst "ReportID=" & first: rows.Bookmark = r.Bookmark: r.Close
 rows!AssignmentRecommendationA = String$(21, "X"): UXGridQueue "AssignmentRecommendationA"
 rows!Trait39 = "1.0": UXGridQueue "Trait39"
 If UXGridDrain() Then Err.Raise vbObjectError + 7923, , "Invalid queue accepted"
 Set rows = Forms!frmUXComparison!uxTraitGrid.Form
 Set r = rows.RecordsetClone: r.FindFirst "ReportID=" & first: rows.Bookmark = r.Bookmark: r.Close
 If rows!Trait39 <> "1.0" Or Not IsNull(DLookup("Trait39", "Report", "ReportID=" & first)) Or UXGridPendingCount() <> 2 Then Err.Raise vbObjectError + 7924, , "Failed queue did not preserve unsaved values separately from saved values"
 rows!AssignmentRecommendationA = "VALID": UXGridQueue "AssignmentRecommendationA"
 rows!Trait39 = "1.0": UXGridQueue "Trait39"
 If Not UXGridDrain() Then Err.Raise vbObjectError + 7925, , "Retry failed"
 TestTraitQueue = "PASS: E6 1.0 marks, row identity, coalescing, failed-queue preservation and retry": Exit Function
Failed:
 TestTraitQueue = "FAIL: " & Err.Number & " " & Err.Description
End Function
