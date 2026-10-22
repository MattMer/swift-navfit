Attribute VB_Name = "modContributorStartup"
Option Compare Binary
Option Explicit
Private compatibilityContributor As Long
Private lastStartupFailure As String

Public Function ContributorStartupFailure() As String
 ContributorStartupFailure = lastStartupFailure
End Function

'Session selection is separate from retained Workspace.ActiveProfileID.
'TempVars survive a VBA reset; path and workspace identity prevent cross-file reuse.
Public Sub ResetActiveContributor()
 On Error Resume Next
 TempVars.Remove "SwiftContributorID"
 TempVars.Remove "SwiftContributorPath"
 TempVars.Remove "SwiftContributorWorkspace"
 Err.Clear
 On Error GoTo 0
End Sub

Public Sub BindActiveContributor(ByVal profile As Long)
 ResetActiveContributor
 TempVars.Add "SwiftContributorID", profile
 TempVars.Add "SwiftContributorPath", CurrentProject.FullName
 TempVars.Add "SwiftContributorWorkspace", CStr(Nz(DLookup("WorkspaceGuid", "Workspace", "WorkspaceID=1"), ""))
End Sub

Public Function HasActiveContributor() As Boolean
 Dim profile As Long, identity As String
 On Error GoTo Missing
 profile = CLng(TempVars("SwiftContributorID").Value)
 If profile < 1 Then Exit Function
 If StrComp(CStr(TempVars("SwiftContributorPath").Value), CurrentProject.FullName, vbTextCompare) <> 0 Then Exit Function
 identity = CStr(Nz(DLookup("WorkspaceGuid", "Workspace", "WorkspaceID=1"), ""))
 If Len(identity) = 0 Or identity <> CStr(TempVars("SwiftContributorWorkspace").Value) Then Exit Function
 If Nz(DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1"), 0) <> profile Then Exit Function
 HasActiveContributor = (DCount("*", "UserProfile", "UserProfileID=" & profile) = 1)
Missing:
End Function

Public Function ActiveContributorID() As Long
 'Historical directory backfill retains its source actor; it never confirms a session.
 If compatibilityContributor > 0 Then ActiveContributorID = compatibilityContributor: Exit Function
 If Not HasActiveContributor() Then Err.Raise vbObjectError + 8900, "ActiveContributorID", "Choose your contributor for this session before making changes."
 ActiveContributorID = CLng(TempVars("SwiftContributorID").Value)
End Function

Public Function MigrateDirectoryCompatibility() As Long
 Dim failure As Long, message As String
 On Error GoTo Failed
 compatibilityContributor = Nz(MajorScalar("SELECT ActiveProfileID FROM Workspace WHERE WorkspaceID=1"), 0)
 MigrateDirectoryCompatibility = SyncDirectoryFromReports()
 compatibilityContributor = 0
 Exit Function
Failed:
 failure = Err.Number: message = Err.Description
 compatibilityContributor = 0
 Err.Raise failure, "MigrateDirectoryCompatibility", message
End Function

Public Function ActiveContributorCaption() As String
 ActiveContributorCaption = "Choose a contributor"
 If HasActiveContributor() Then ActiveContributorCaption = Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & ActiveContributorID()), "")
End Function

Public Function LoadContributorStartup() As Boolean
 With Forms!frmContributorStart
  !cboContributor.RowSource = "SELECT UserProfileID,DisplayName FROM UserProfile ORDER BY DisplayName,UserProfileID"
  !cboContributor = Null
  !txtNewContributor = Null
  !lblStatus.Caption = "Choose your name for this launch, or add a contributor."
 End With
 LoadContributorStartup = True
End Function

Public Function ShowContributorStartup() As Boolean
 If DCount("*", "Workspace", "WorkspaceID=1") = 0 Then
  DoCmd.OpenForm "frmSetup"
 Else
  DoCmd.OpenForm "frmContributorStart"
  Forms!frmContributorStart.SetFocus
 End If
 ShowContributorStartup = True
End Function

Public Function AcceptContributorStartup(Optional ByVal createNew As Boolean = False) As Boolean
 Dim profile As Long, name As String, w As DAO.Workspace, db As DAO.Database, rows As DAO.Recordset
 Dim trans As Boolean, message As String, stage As String
 On Error GoTo Failed
 lastStartupFailure = ""
 stage = "Validate contributor"
 If createNew Then
  RequireContributorSwitch
  name = Trim$(Nz(Forms!frmContributorStart!txtNewContributor, ""))
  If Len(name) = 0 Or Len(name) > 120 Then Err.Raise 5, , "Enter a contributor name of 1 to 120 characters."
  If DCount("*", "UserProfile", "DisplayName=" & QuoteSQL(name)) > 0 Then Err.Raise 5, , "That contributor already exists. Choose the existing name."
  stage = "Create contributor"
  Set db = CurrentDb: Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
  Set rows = db.OpenRecordset("UserProfile", dbOpenDynaset)
  rows.AddNew: rows!ProfileGuid = NewIdentity(): rows!DisplayName = name: rows.Update
  rows.Bookmark = rows.LastModified: profile = rows!UserProfileID: rows.Close: Set rows = Nothing
  'Domain lookups do not see this transaction's new profile until commit.
  'Keep profile creation and retained selection in the same owned transaction.
  db.Execute "UPDATE Workspace SET ActiveProfileID=" & profile & " WHERE WorkspaceID=1", dbFailOnError
  stage = "Commit contributor"
  w.CommitTrans: trans = False
  BindActiveContributor profile
 Else
  profile = CLng(Nz(Forms!frmContributorStart!cboContributor, 0))
  stage = "Select contributor"
  Call SwitchContributor(profile)
 End If
 stage = "Open workspace"
 DoCmd.Close acForm, "frmContributorStart", acSaveNo
 DoCmd.OpenForm "frmHome"
 UXLoadContext
 AcceptContributorStartup = True: Exit Function
Failed:
 message = stage & ": " & Err.Description
 lastStartupFailure = message
 If trans Then
  On Error Resume Next
  If Not rows Is Nothing Then rows.Close
  w.Rollback
  ResetActiveContributor
  On Error GoTo 0
 End If
 If Not CurrentProject.AllForms("frmContributorStart").IsLoaded Then DoCmd.OpenForm "frmContributorStart"
 Forms!frmContributorStart!lblStatus.Caption = message
End Function

Public Function CancelContributorStartup() As Boolean
 ResetActiveContributor
 Application.Quit acQuitSaveAll
End Function
