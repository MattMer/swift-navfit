Attribute VB_Name = "modSavedReadiness"
Option Explicit
Public Const SHARED_READINESS_VERSION As String = "1610.10H CH-2 + NAVADMIN 132/26 / applicability-v1"
' Install a new immutable pack; existing definitions and historical results remain intact.
Public Function InstallSharedReadinessPack() As Long
 Dim db As DAO.Database, r As DAO.Recordset, w As DAO.Workspace, id As Long, failure As Long, message As String
 id = Nz(MajorScalar("SELECT PackID FROM RulePack WHERE VersionName=" & QuoteSQL(SHARED_READINESS_VERSION)), 0)
 If id > 0 Then InstallSharedReadinessPack = id: Exit Function
 Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans
 On Error GoTo Failed
 Set r = db.OpenRecordset("RulePack", dbOpenDynaset): r.AddNew
 r!VersionName = SHARED_READINESS_VERSION: r!SourceTitle = "CH-2 / CH-1 calendars; NAVADMIN 132/26 effective 3 June 2026"
 r!SourceURL = "https://www.mynavyhr.navy.mil/References/BUPERS-Instructions/": r!SourceSHA256 = NAVY_RULE_SOURCE_HASH
 r!EffectiveFrom = #12/16/2025#: r!EffectiveTo = Null: r.Update: r.Bookmark = r.LastModified: id = r!PackID: r.Close
 Set r = db.OpenRecordset("RuleDefinition", dbOpenDynaset): r.AddNew
 r!PackID = id: r!RuleCode = "SHARED-READINESS": r!FieldKey = "OfficialName": r!CheckKind = "SharedReadiness"
 r!SourceLocator = "Each finding retains its own policy or native compatibility reference"
 r!Description = "Versioned applicability: CH-1 calendars from 16 December 2025; general CH-2 policy from 26 May 2026; NAVADMIN 132/26 exemptions from 3 June 2026. Unsupported dates require review. Saved revisions only.": r!Severity = "Review"
 r.Update: r.Close: w.CommitTrans: InstallSharedReadinessPack = id: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modSavedReadiness", "InstallSharedReadinessPack", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
 failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description")): SwiftRollback w, swiftPrimary
 If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
 On Error GoTo 0: SwiftRaise swiftPrimary
End Function
' Persist the same shared findings against one saved revision, including full references.
Public Function RunSharedReadiness(ByVal reportID As Long, ByVal packID As Long) As Long
 Dim db As DAO.Database, w As DAO.Workspace, findings As Collection, item As Object, r As DAO.Recordset
 Dim revision As Long, rule As Long, id As Long, state As String, failure As Long, message As String, recorded As Long
 If Nz(MajorScalar("SELECT VersionName FROM RulePack WHERE PackID=" & packID), "") <> SHARED_READINESS_VERSION Then Err.Raise vbObjectError + 8403, "modSavedReadiness.RunSharedReadiness", "Historical readiness pack retained for evidence only. Install the current pack to create a new validation run."
 If Nz(MajorScalar("SELECT SourceSHA256 FROM RulePack WHERE PackID=" & packID), "") <> NAVY_RULE_SOURCE_HASH Then Err.Raise vbObjectError + 8402, "modSavedReadiness.RunSharedReadiness", "Shared readiness source fingerprint differs. Restore the verified rule pack."
 revision = CurrentRevision(reportID): Set findings = ValidateLiveFields(SavedFields(reportID), reportID)
 rule = Nz(MajorScalar("SELECT RuleID FROM RuleDefinition WHERE PackID=" & packID & " AND CheckKind='SharedReadiness'"), 0)
 If rule = 0 Then Err.Raise vbObjectError + 8400, "modSavedReadiness.RunSharedReadiness", "Shared readiness rule definition is missing."
 Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans
 On Error GoTo Failed
 If CurrentRevision(reportID) <> revision Then Err.Raise vbObjectError + 8401, "modSavedReadiness.RunSharedReadiness", "Report changed. Run saved readiness again."
 Set r = db.OpenRecordset("ValidationRun", dbOpenDynaset): r.AddNew
 r!ReportID = reportID: r!RevisionID = revision: r!PackID = packID: r!RunAt = Now(): r.Update: r.Bookmark = r.LastModified: id = r!ValidationRunID: r.Close
 Set r = db.OpenRecordset("ValidationFinding", dbOpenDynaset)
 'Information findings (print line fit) are not readiness results.
 For Each item In findings
  If item("Severity") <> "Information" Then
   state = "NeedsReview": If item("Severity") <> "Needs review" Then state = "Fail"
   r.AddNew: r!ValidationRunID = id: r!RuleID = rule: r!FindingState = state
   r!Message = item("Code") & " | " & item("Label") & " | " & item("Severity") & ": " & item("Message") & vbCrLf & item("Reference"): r.Update
   recorded = recorded + 1
  End If
 Next
 If recorded = 0 Then
  r.AddNew: r!ValidationRunID = id: r!RuleID = rule: r!FindingState = "Pass": r!Message = "No findings from implemented shared checks. This is not certification of Navy readiness.": r.Update
 End If
 r.Close: w.CommitTrans: RunSharedReadiness = id: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modSavedReadiness", "RunSharedReadiness", "Execute", Err.Number, Err.Description, Err.Source, "Unknown")
 failure = CLng(swiftPrimary("Number")): message = CStr(swiftPrimary("Description")): SwiftRollback w, swiftPrimary
 If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
 On Error GoTo 0: SwiftRaise swiftPrimary
End Function
