Attribute VB_Name = "modRSCompletion"
Option Explicit
'Local completion entries append immutable baselines. Policy markers retain the
'group contribution across reopen/archive; later completion posts only its delta.
Private Const LocalPrefix As String = "Swift local RSCA; parent="
Public Function RSLatestBaseline(ByVal seniorID As Long, ByVal grade As String) As Long
 Dim r As DAO.Recordset
 Set r = CurrentDb.OpenRecordset("SELECT TOP 1 BaselineID FROM RSCABaseline WHERE SeniorID=" & seniorID & " AND Paygrade=" & QuoteSQL(grade) & " ORDER BY AsOfDate DESC,BaselineID DESC", dbOpenSnapshot)
 If Not r.EOF Then RSLatestBaseline = r!BaselineID
 r.Close
End Function
Public Function RSBaselineIncludes(ByVal baselineID As Long, ByVal ancestor As Long) As Boolean
 Dim note As String, parent As Long
 Do While baselineID > 0 And ancestor > 0
  If baselineID = ancestor Then RSBaselineIncludes = True: Exit Function
  parent = Nz(DLookup("ParentBaselineID", "RSCABaseline", "BaselineID=" & baselineID), 0)
  If parent = 0 Then Exit Function
  If parent >= baselineID Then Err.Raise vbObjectError + 4850, , "RSCA baseline ancestry is invalid. Inspect history before continuing."
  baselineID = parent
 Loop
End Function
Public Function RSCalculationBaseline(ByVal groupID As Long, ByVal selected As Long) As Long
 Dim seniorID As Long, grade As String, latest As Long
 RSCalculationBaseline = selected
 If selected = 0 Then Exit Function
 seniorID = Nz(DLookup("SeniorID", "GroupRSCABaseline", "GroupID=" & groupID), 0)
 grade = Nz(DLookup("Paygrade", "SummaryGroup", "GroupID=" & groupID), "")
 latest = RSLatestBaseline(seniorID, grade)
 If RSBaselineIncludes(latest, selected) Then RSCalculationBaseline = latest
End Function
Public Sub RSExcludePostedGroup(ByVal groupID As Long, ByVal baselineID As Long, ByRef total As Double, ByRef count As Long)
 Dim posted As Long
 posted = Val(GroupPolicyValue(groupID, "RSCAPostedBaseline"))
 If posted = 0 Then Exit Sub
 If Not RSBaselineIncludes(baselineID, posted) Then Err.Raise vbObjectError + 4851, , "This group was already counted. The selected history does not contain its local completion entry. Reconcile the BOL observation with local completion history before continuing."
 total = total - CDbl(GroupPolicyValue(groupID, "RSCAGroupSum"))
 count = count - CLng(GroupPolicyValue(groupID, "RSCAGroupCount"))
 If count < 0 Or total < 0 Then Err.Raise vbObjectError + 4852, , "Recorded completion exceeds historical totals. Inspect RSCA history."
End Sub
'Called inside SetGroupWorkflow's transaction: errors roll back status, baseline,
'operation history and markers together. No report revision is changed.
Public Sub RSPostGroupCompletion(ByVal groupID As Long, ByVal policy As Object)
 Dim p As Object, r As DAO.Recordset, groups As DAO.Recordset, key As Variant
 Dim ids As String, id As Variant, latest As Long, posted As Long, baseline As Long
 Dim grade As String, total As Double, count As Long, exact As Variant, asOf As Date, average As Variant
 If CurrentProject.AllForms("frmEditor").IsLoaded Then
  If DCount("*", "GroupMember", "GroupID=" & groupID & " AND ReportID=" & Nz(Forms!frmEditor!txtReportID, 0)) > 0 Then
   If EditorHasPending() Then Err.Raise vbObjectError + 4853, , "Save or discard pending member edits before completing this group."
  End If
 End If
 Set p = RSGroupPSG(groupID)
 If Not p.Exists("GroupCount") Then Err.Raise vbObjectError + 4854, , "Cannot complete RSCA accounting: " & p("Reason")
 grade = CStr(DLookup("Paygrade", "SummaryGroup", "GroupID=" & groupID))
 latest = RSLatestBaseline(CLng(p("SeniorID")), grade)
 If latest <> CLng(p("BaselineID")) Then Err.Raise vbObjectError + 4855, , "Select the latest matching paygrade baseline before completing this group. Earlier observations remain in history."
 Set r = CurrentDb.OpenRecordset("SELECT r.* FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & groupID & " ORDER BY r.ReportID", dbOpenSnapshot): ids = "|"
 Do While Not r.EOF
  If Nz(r!ObservationStatus, "") <> "NOB" Then
   If Nz(r!ObservationStatus, "") <> "Observed" Then Err.Raise vbObjectError + 4856, , "Choose Observed or whole-report NOB for every member."
   For Each key In TraitKeys()
    Select Case Nz(r.Fields(CStr(key)).Value, "")
     Case "1.0", "2.0", "3.0", "4.0", "5.0", "NOB"
     Case Else: Err.Raise vbObjectError + 4856, , "Finish all observed member traits before completing this group. Explicit NOB traits are excluded."
    End Select
   Next
   ids = ids & r!ReportID & "|"
  End If
  r.MoveNext
 Loop
 r.Close
 For Each id In Split(Mid$(ids, 2), "|")
  If Len(id) > 0 Then
   If MajorScalar("SELECT Count(*) FROM RSCountedReport AS c INNER JOIN GroupCompletion AS p ON c.CompletionID=p.CompletionID WHERE c.ReportID=" & CLng(id) & " AND p.GroupID<>" & groupID) > 0 Then Err.Raise vbObjectError + 4857, , "A report is already counted through another group. Reconcile its accounting before completion."
  End If
 Next
 posted = Val(GroupPolicyValue(groupID, "RSCAPostedBaseline"))
 If posted > 0 Then
  If GroupPolicyValue(groupID, "RSCAReportIDs") = ids And CDbl(GroupPolicyValue(groupID, "RSCAGroupSum")) = CDbl(p("GroupSum")) And CLng(GroupPolicyValue(groupID, "RSCAGroupCount")) = CLng(p("GroupCount")) And RSPostingRevisionsMatch(groupID) Then Exit Sub
 End If
 total = CDbl(p("HistoricalSum")) + CDbl(p("GroupSum"))
 count = CLng(p("HistoricalCount")) + CLng(p("GroupCount"))
 exact = Null
 If Not p("Estimated") Then exact = total
 asOf = Date
 If CDate(DLookup("AsOfDate", "RSCABaseline", "BaselineID=" & latest)) > asOf Then asOf = CDate(DLookup("AsOfDate", "RSCABaseline", "BaselineID=" & latest))
 average = Null
 If count > 0 Then average = total / count
 baseline = RSAddBaseline(CLng(p("SeniorID")), grade, asOf, average, count, exact, LocalPrefix & latest & "; group=" & groupID & "; local completion, not BOL verification")
 CurrentDb.Execute "UPDATE RSCABaseline SET ParentBaselineID=" & latest & " WHERE BaselineID=" & baseline, dbFailOnError
 Set r = RSGroupRows(groupID)
 RSWritePosting groupID, CLng(p("SeniorID")), latest, baseline, CDbl(p("GroupSum")), CLng(p("GroupCount")), r
 r.Close
End Sub
