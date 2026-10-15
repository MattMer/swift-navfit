Attribute VB_Name = "modSavedReadiness"
Option Explicit
' Install a new immutable pack; existing definitions and historical results remain intact.
Public Function InstallSharedReadinessPack() As Long
 Dim db As DAO.Database, r As DAO.Recordset, w As DAO.Workspace, id As Long, failure As Long, message As String
 id = Nz(MajorScalar("SELECT PackID FROM RulePack WHERE VersionName='1610.10H CH-2 / shared-readiness-r110'"), 0)
 If id > 0 Then InstallSharedReadinessPack = id: Exit Function
 Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans
 On Error GoTo Failed
 Set r = db.OpenRecordset("RulePack", dbOpenDynaset): r.AddNew
 r!VersionName = "1610.10H CH-2 / shared-readiness-r110": r!SourceTitle = "BUPERSINST 1610.10H CH-2 / shared report readiness r110"
 r!SourceURL = "https://www.mynavyhr.navy.mil/References/BUPERS-Instructions/": r!SourceSHA256 = NAVY_RULE_SOURCE_HASH
 r!EffectiveFrom = #5/26/2026#: r!EffectiveTo = Null: r.Update: r.Bookmark = r.LastModified: id = r!PackID: r.Close
 Set r = db.OpenRecordset("RuleDefinition", dbOpenDynaset): r.AddNew
 r!PackID = id: r!RuleCode = "SHARED-READINESS": r!FieldKey = "OfficialName": r!CheckKind = "SharedReadiness"
 r!SourceLocator = "Each finding retains its own policy or native compatibility reference"
 r!Description = "Same finding engine as live report and group validation; pending values are not included in saved runs.": r!Severity = "Review"
 r.Update: r.Close: w.CommitTrans: InstallSharedReadinessPack = id: Exit Function
Failed:
 failure = Err.Number: message = Err.Description: On Error Resume Next: w.Rollback
 If Not r Is Nothing Then r.Close
 On Error GoTo 0: Err.Raise failure, "InstallSharedReadinessPack", message
End Function
' Persist the same shared findings against one saved revision, including full references.
Public Function RunSharedReadiness(ByVal reportID As Long, ByVal packID As Long) As Long
 Dim db As DAO.Database, w As DAO.Workspace, findings As Collection, item As Object, r As DAO.Recordset
 Dim revision As Long, rule As Long, id As Long, state As String, failure As Long, message As String
 If Nz(MajorScalar("SELECT SourceSHA256 FROM RulePack WHERE PackID=" & packID), "") <> NAVY_RULE_SOURCE_HASH Then Err.Raise vbObjectError + 8402, , "Shared readiness source fingerprint differs. Restore the verified rule pack."
 revision = CurrentRevision(reportID): Set findings = ValidateLiveFields(SavedFields(reportID), reportID)
 rule = Nz(MajorScalar("SELECT RuleID FROM RuleDefinition WHERE PackID=" & packID & " AND CheckKind='SharedReadiness'"), 0)
 If rule = 0 Then Err.Raise vbObjectError + 8400, , "Shared readiness rule definition is missing."
 Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans
 On Error GoTo Failed
 If CurrentRevision(reportID) <> revision Then Err.Raise vbObjectError + 8401, , "Report changed. Run saved readiness again."
 Set r = db.OpenRecordset("ValidationRun", dbOpenDynaset): r.AddNew
 r!ReportID = reportID: r!RevisionID = revision: r!PackID = packID: r!RunAt = Now(): r.Update: r.Bookmark = r.LastModified: id = r!ValidationRunID: r.Close
 Set r = db.OpenRecordset("ValidationFinding", dbOpenDynaset)
 For Each item In findings
  state = "NeedsReview": If item("Severity") <> "Needs review" Then state = "Fail"
  r.AddNew: r!ValidationRunID = id: r!RuleID = rule: r!FindingState = state
  r!Message = item("Code") & " | " & item("Label") & " | " & item("Severity") & ": " & item("Message") & vbCrLf & item("Reference"): r.Update
 Next
 If findings.Count = 0 Then
  r.AddNew: r!ValidationRunID = id: r!RuleID = rule: r!FindingState = "Pass": r!Message = "No findings from implemented shared checks. This is not certification of Navy readiness.": r.Update
 End If
 r.Close: w.CommitTrans: RunSharedReadiness = id: Exit Function
Failed:
 failure = Err.Number: message = Err.Description: On Error Resume Next: w.Rollback
 If Not r Is Nothing Then r.Close
 On Error GoTo 0: Err.Raise failure, "RunSharedReadiness", message
End Function
