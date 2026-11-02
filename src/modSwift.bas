Attribute VB_Name = "modSwift"
Option Compare Database
Option Explicit

Private Type SwiftGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef value As SwiftGuid) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef value As SwiftGuid, ByVal buffer As LongPtr, ByVal capacity As Long) As Long
Public Const APP_VERSION As String = "0.52.0-dev"
Public Const SCHEMA_VERSION As Long = 38

Public Function NewIdentity() As String
    Dim value As SwiftGuid, buffer As String, result As Long
    If CoCreateGuid(value) <> 0 Then Err.Raise vbObjectError + 100, "modSwift.NewIdentity", "Could not create a stable identity."
    buffer = String$(39, vbNullChar)
    result = StringFromGUID2(value, StrPtr(buffer), 39)
    If result = 0 Then Err.Raise vbObjectError + 101, "modSwift.NewIdentity", "Could not format identity."
    NewIdentity = Left$(buffer, result - 1)
End Function

' Bind DAO query parameters using an owned database reference kept alive until
' execution completes; return RecordsAffected and let callers handle transactional failures.
Private Function ExecuteParams(ByVal statement As String, ParamArray values() As Variant) As Long
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim q As DAO.QueryDef, i As Long
    Set q = ownedDB.CreateQueryDef("", statement)
    For i = 0 To UBound(values)
        q.Parameters(i).Value = values(i)
    Next
    q.Execute dbFailOnError
    ExecuteParams = q.RecordsAffected
    q.Close
End Function

Private Function ProfileID() As Long
    ProfileID = ActiveContributorID()
End Function

Public Function ConfigureWorkspace(ByVal commandName As String, ByVal displayName As String) As Boolean
    Dim w As DAO.Workspace, db As DAO.Database, r As DAO.Recordset, pid As Long
    commandName = Trim$(commandName): displayName = Trim$(displayName)
    If Len(commandName) = 0 Or Len(displayName) = 0 Then Err.Raise vbObjectError + 103, "modSwift.ConfigureWorkspace", "Enter a command name and contributor name."
    If Len(commandName) > 120 Or Len(displayName) > 120 Then Err.Raise vbObjectError + 104, "modSwift.ConfigureWorkspace", "Names must be 120 characters or fewer."
    If DCount("*", "Workspace", "WorkspaceID=1") > 0 Then Call ActiveContributorID
    Set w = DBEngine.Workspaces(0): Set db = CurrentDb
    On Error GoTo Failed
    w.BeginTrans
    pid = Nz(DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1"), 0)
    If pid = 0 Then
        Set r = db.OpenRecordset("UserProfile", dbOpenDynaset)
        r.AddNew: r!ProfileGuid = NewIdentity(): r!DisplayName = displayName: r.Update
        r.Bookmark = r.LastModified: pid = r!UserProfileID: r.Close
        Set r = db.OpenRecordset("Workspace", dbOpenDynaset)
        r.AddNew: r!WorkspaceID = 1: r!WorkspaceGuid = NewIdentity()
    Else
        ExecuteParams "PARAMETERS p Text(120), id Long; UPDATE UserProfile SET DisplayName=p WHERE UserProfileID=id", displayName, pid
        Set r = db.OpenRecordset("SELECT * FROM Workspace WHERE WorkspaceID=1", dbOpenDynaset)
        r.Edit
    End If
    r!CommandName = commandName: r!ActiveProfileID = pid: r.Update: r.Close
    w.CommitTrans
    BindActiveContributor pid
    ConfigureWorkspace = True
    Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modSwift", "ConfigureWorkspace", "Execute", Err.number, Err.description, Err.source, "Unknown")
    SwiftRollback w, swiftPrimary
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function

'Callers pass one append-only recordset per commit. A full dynaset reopened per field
'costs time proportional to the table inside a transaction with pending inserts.
Private Sub Snapshot(ByVal r As DAO.Recordset, ByVal revisionID As Long, ByVal fieldName As String, ByVal value As Variant)
    r.AddNew: r!RevisionID = revisionID: r!FieldKey = fieldName
    r!ValueText = value: r!IsNullValue = IsNull(value): r.Update
End Sub

Private Sub AppendChange(ByVal r As DAO.Recordset, ByVal revisionID As Long, ByVal fieldName As String, ByVal oldValue As Variant, ByVal newValue As Variant)
    r.AddNew: r!RevisionID = revisionID: r!FieldKey = fieldName
    r!OldValue = oldValue: r!NewValue = newValue
    r!OldIsNull = IsNull(oldValue): r!NewIsNull = IsNull(newValue)
    r.Update
End Sub

Private Function FieldText(ByVal value As Variant) As Variant
    If IsNull(value) Then
        FieldText = Null
    ElseIf VarType(value) = vbDate Then
        FieldText = Format$(value, "yyyy-mm-dd")
    Else
        FieldText = CStr(value)
    End If
End Function

' Validate and atomically commit report values, their revision snapshot and audit changes.
' Keep the report write cursor alive through revision-pointer assignment to avoid DAO
' lock failures; invalidate editor context after a successful local save.
Public Function CommitDraft(ByVal reportID As Long, ByVal formKind As String, ByVal title As String, ByVal memberName As String, ByVal startDate As Variant, ByVal endDate As Variant, ByVal notes As String, Optional ByVal reason As String = "Draft save", Optional ByVal officialComments As Variant, Optional ByVal administration As Variant, Optional ByVal traits As Variant, Optional ByVal assessment As Variant, Optional ByVal promotionStatus As Variant, Optional ByVal reportingSenior As Variant, Optional ByVal reportContext As Variant, Optional ByVal memberDetails As Variant, Optional ByVal serviceDetails As Variant, _
    Optional ByVal decisionDetails As Variant, Optional ByVal reportPeople As Variant, Optional ByVal partyNames As Variant, Optional ByVal commandEmployment As Variant, Optional ByVal additionalFields As Variant, Optional ByVal auditOperationID As Long = 0) As Long
    Dim w As DAO.Workspace, db As DAO.Database, r As DAO.Recordset, rev As DAO.Recordset, snapshotRows As DAO.Recordset, changeRows As DAO.Recordset
    Dim keys As Variant, previous() As Variant, current() As Variant, i As Long
    Dim revisionID As Long, seq As Long, parentID As Variant, pid As Long, changed As Boolean, changedKeys As String
    If formKind <> "EVAL" And formKind <> "CHIEFEVAL" And formKind <> "FITREP" Then Err.Raise vbObjectError + 105, "modSwift.CommitDraft", "Choose EVAL, CHIEFEVAL, or FITREP."
    If Len(title) > UXFieldLimit("Title") Or Len(memberName) > UXFieldLimit("MemberName") Then Err.Raise vbObjectError + 106, "modSwift.CommitDraft", "Title or member name exceeds local draft capacity."
    If Not IsNull(startDate) Then
        If Not IsDate(startDate) Then Err.Raise vbObjectError + 107, "modSwift.CommitDraft", "Start date is not a valid date."
        startDate = DateValue(startDate)
    End If
    If Not IsNull(endDate) Then
        If Not IsDate(endDate) Then Err.Raise vbObjectError + 108, "modSwift.CommitDraft", "End date is not a valid date."
        endDate = DateValue(endDate)
    End If
    pid = ProfileID()
    If reportID > 0 Then AssertReportCache reportID
    keys = DraftFieldKeys()
    ReDim previous(0 To UBound(keys)): ReDim current(0 To UBound(keys))
    For i = 0 To UBound(keys)
        previous(i) = Null
        current(i) = Null
    Next
    current(0) = formKind: current(1) = title: current(2) = memberName
    current(3) = FieldText(startDate): current(4) = FieldText(endDate): current(5) = notes
    If IsMissing(officialComments) Then
        If reportID = 0 Then current(6) = Null Else current(6) = StoredOfficialNarrative(reportID)
    Else
        If Not IsNull(officialComments) And VarType(officialComments) <> vbString Then Err.Raise vbObjectError + 117, "modSwift.CommitDraft", "Official narrative must be text or null."
        current(6) = officialComments
    End If
    If Not IsMissing(commandEmployment) Then
        If Not IsNull(commandEmployment) And VarType(commandEmployment) <> vbString Then Err.Raise vbObjectError + 118, "modSwift.CommitDraft", "Block 28 must be text or null."
        current(78) = commandEmployment
    End If
    If Not IsMissing(administration) Then
        administration = ValidateAdministration(administration)
        For i = 0 To 4: current(7 + i) = administration(i): Next
    End If
    If Not IsMissing(traits) Then
        traits = ValidateTraits(traits)
        For i = 0 To 6: current(12 + i) = traits(i): Next
    End If
    If Not IsMissing(assessment) Then
        assessment = ValidateAssessment(assessment)
        current(19) = assessment(0): current(20) = assessment(1)
    End If
    If Not IsMissing(serviceDetails) Then
        serviceDetails = ValidateServiceDetails(serviceDetails)
        For i = 0 To 5: current(43 + i) = serviceDetails(i): Next
    End If
    If Not IsMissing(decisionDetails) Then
        decisionDetails = ValidateDecisionDetails(decisionDetails)
        For i = 0 To 5: current(49 + i) = decisionDetails(i): Next
    End If
    If Not IsMissing(reportPeople) Then
        reportPeople = ValidateReportPeople(reportPeople)
        For i = 0 To 12: current(55 + i) = reportPeople(i): Next
    End If
    If Not IsMissing(partyNames) Then
        partyNames = ValidatePartyNames(partyNames)
        For i = 0 To 9: current(68 + i) = partyNames(i): Next
    End If
    If Not IsMissing(memberDetails) Then
        memberDetails = ValidateMemberDetails(memberDetails)
        For i = 0 To 6: current(36 + i) = memberDetails(i): Next
    End If
    If Not IsMissing(reportContext) Then
        reportContext = ValidateReportContext(reportContext)
        For i = 0 To 8: current(27 + i) = reportContext(i): Next
    End If
    If Not IsMissing(reportingSenior) Then
        reportingSenior = ValidateReportingSenior(reportingSenior)
        For i = 0 To 4: current(22 + i) = reportingSenior(i): Next
    End If
    If Not IsMissing(promotionStatus) Then current(21) = ValidatePromotionStatus(promotionStatus)
    Set w = DBEngine.Workspaces(0): Set db = CurrentDb
    On Error GoTo Failed
    w.BeginTrans
    If reportID = 0 Then
        Set r = db.OpenRecordset("Report", dbOpenDynaset)
        r.AddNew: r!ReportGuid = NewIdentity(): r!CreatedAt = Now(): r!WorkflowState = "Draft"
        r!IsArchived = False: changed = True: seq = 1: parentID = Null
    Else
        Set r = db.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & reportID, dbOpenDynaset)
        If r.EOF Then Err.Raise vbObjectError + 109, "modSwift.CommitDraft", "Report does not exist."
        If r!IsArchived Then Err.Raise vbObjectError + 110, "modSwift.CommitDraft", "Recover archived reports before editing."
        If r!WorkflowState <> "Draft" Then Err.Raise vbObjectError + 111, "modSwift.CommitDraft", "Reopen locked reports before editing."
        If r!formKind <> formKind Then Err.Raise vbObjectError + 112, "modSwift.CommitDraft", "Report type cannot change after creation."
        If IsMissing(administration) Then
            For i = 7 To 11: current(i) = r.Fields(keys(i)).Value: Next
        End If
        If IsMissing(traits) Then
            For i = 12 To 18: current(i) = r.Fields(keys(i)).Value: Next
        End If
        If IsMissing(assessment) Then current(19) = r!ObservationStatus.Value: current(20) = r!PromotionRecommendation.Value
        If IsMissing(promotionStatus) Then current(21) = r!PromotionStatus.Value
        If IsMissing(reportingSenior) Then
            For i = 22 To 26: current(i) = r.Fields(keys(i)).Value: Next
        End If
        If IsMissing(reportContext) Then
            For i = 27 To 35: current(i) = r.Fields(keys(i)).Value: Next
        End If
        If IsMissing(memberDetails) Then
            For i = 36 To 42: current(i) = r.Fields(keys(i)).Value: Next
        End If
        If IsMissing(serviceDetails) Then
            For i = 43 To 48: current(i) = r.Fields(keys(i)).Value: Next
        End If
        If IsMissing(decisionDetails) Then
            For i = 49 To 54: current(i) = r.Fields(keys(i)).Value: Next
        End If
        If IsMissing(reportPeople) Then
            For i = 55 To 67: current(i) = r.Fields(keys(i)).Value: Next
        End If
        If IsMissing(partyNames) Then
            For i = 68 To 77: current(i) = r.Fields(keys(i)).Value: Next
        End If
        If IsMissing(commandEmployment) Then current(78) = StoredBlock28(reportID)
        For i = 79 To UBound(keys): current(i) = r.Fields(keys(i)).Value: Next
        If Not IsMissing(additionalFields) Then
            For i = 79 To UBound(keys): current(i) = additionalFields(CStr(keys(i))): Next
        End If
        For i = 79 To UBound(keys)
            ValidateCatalogueField CStr(keys(i)), current(i): current(i) = FieldText(current(i))
        Next
        For i = 0 To UBound(keys)
            If i = 78 Then previous(i) = FieldText(StoredBlock28(reportID)) Else previous(i) = FieldText(r.Fields(keys(i)).Value)
            If IsNull(previous(i)) <> IsNull(current(i)) Or StrComp(Nz(previous(i), ""), Nz(current(i), ""), vbBinaryCompare) <> 0 Then changed = True
        Next
        If Not changed Then
            r.Close: SyncReportDirectory reportID: w.CommitTrans: Call UXEditorContextChanged(): CommitDraft = reportID: Exit Function
        End If
        'Under review only the holder saves; Members-grid saves are logged instead.
        RouteSaveGuard reportID
        parentID = r!CurrentRevisionID
        Set rev = db.OpenRecordset("SELECT Max(RevisionNumber) FROM ReportRevision WHERE ReportID=" & reportID, dbOpenSnapshot)
        seq = Nz(rev.Fields(0).Value, 0) + 1: rev.Close: Set rev = Nothing
        r.Edit
    End If
    If reportID = 0 And Not IsMissing(additionalFields) Then
        For i = 79 To UBound(keys): current(i) = additionalFields(CStr(keys(i))): Next
    End If
    For i = 79 To UBound(keys)
        ValidateCatalogueField CStr(keys(i)), current(i): current(i) = FieldText(current(i))
    Next
    r!FormKind = formKind: r!Title = title: r!MemberName = memberName
    r!PeriodStart = startDate: r!PeriodEnd = endDate: r!DraftNotes = notes: r!OfficialComments = current(6): r!ModifiedAt = Now()
    For i = 7 To UBound(keys): r.Fields(keys(i)) = current(i): Next
    r.Update: r.Bookmark = r.LastModified: reportID = r!ReportID
    Set rev = db.OpenRecordset("ReportRevision", dbOpenDynaset)
    rev.AddNew: rev!ReportID = reportID: rev!RevisionNumber = seq: rev!ParentRevisionID = parentID
    rev!AuthorProfileID = pid: rev!CommittedAt = Now(): rev!Reason = Left$(reason, 255): rev!SnapshotSchema = SCHEMA_VERSION
    rev!FieldDefinitionVersion = FIELD_DEFINITION_VERSION
    rev.Update: rev.Bookmark = rev.LastModified: revisionID = rev!RevisionID: rev.Close: Set rev = Nothing
    Set snapshotRows = db.OpenRecordset("RevisionValue", dbOpenDynaset, dbAppendOnly)
    Set changeRows = db.OpenRecordset("RevisionChange", dbOpenDynaset, dbAppendOnly)
    For i = 0 To UBound(keys)
        Snapshot snapshotRows, revisionID, CStr(keys(i)), current(i)
        'A first save starts from Null, so fields left empty log nothing (L03); the
        'snapshot above still records every field.
        If IsNull(previous(i)) <> IsNull(current(i)) Or StrComp(Nz(previous(i), ""), Nz(current(i), ""), vbBinaryCompare) <> 0 Then
            AppendChange changeRows, revisionID, CStr(keys(i)), previous(i), current(i)
            If Len(changedKeys) > 0 Then changedKeys = changedKeys & "|"
            changedKeys = changedKeys & keys(i)
        End If
    Next
    snapshotRows.Close: Set snapshotRows = Nothing
    changeRows.Close: Set changeRows = Nothing
    'Keep the report write cursor alive through revision linking.
    'A second SQL update can contend with its uncommitted record lock.
    r.Edit: r!CurrentRevisionID = revisionID: r.Update
    r.Close: Set r = Nothing
    SelectReportHead reportID, revisionID
    WriteRevisionFacts revisionID
    AssertReportCache reportID
    AuditRecord "Revision", revisionID, auditOperationID
    If Not IsEmpty(parentID) And Not IsNull(parentID) Then RouteRecordSave db, reportID, revisionID, changedKeys
    SyncReportDirectory reportID
    w.CommitTrans
    Call UXEditorContextChanged(): CommitDraft = reportID
    Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modSwift", "CommitDraft", "Execute", Err.number, Err.description, Err.source, "Unknown")
    Dim failure As Long, explanation As String
    failure = CLng(swiftPrimary("Number")): explanation = CStr(swiftPrimary("Description"))
    SwiftRollback w, swiftPrimary
    If Not r Is Nothing Then SwiftClose r, swiftPrimary, "r"
    If Not rev Is Nothing Then SwiftClose rev, swiftPrimary, "rev"
    If Not snapshotRows Is Nothing Then SwiftClose snapshotRows, swiftPrimary, "snapshotRows"
    If Not changeRows Is Nothing Then SwiftClose changeRows, swiftPrimary, "changeRows"
    On Error GoTo 0
    SwiftRaise swiftPrimary
End Function

Public Function ArchiveDraft(ByVal reportID As Long, ByVal archived As Boolean, Optional ByVal auditOperationID As Long = 0) As Boolean
    Dim w As DAO.Workspace, pid As Long
    pid = ProfileID()
    If DCount("*", "Report", "ReportID=" & reportID) = 0 Then Err.Raise vbObjectError + 113, "modSwift.ArchiveDraft", "Report does not exist."
    Set w = DBEngine.Workspaces(0)
    On Error GoTo Failed
    w.BeginTrans
    ExecuteParams "PARAMETERS flag Bit, id Long; UPDATE Report SET IsArchived=flag, ModifiedAt=Now() WHERE ReportID=id", archived, reportID
    Dim eventRow As DAO.Recordset
    Set eventRow = CurrentDb.OpenRecordset("ReportEvent", dbOpenDynaset)
    eventRow.AddNew: eventRow!ReportID = reportID: eventRow!AuthorProfileID = pid: eventRow!EventAt = Now(): eventRow!EventKind = IIf(archived, "Archived", "Recovered")
    eventRow.Update: eventRow.Bookmark = eventRow.LastModified
    AuditRecord "Report event", CLng(eventRow!ReportEventID), auditOperationID: eventRow.Close
    w.CommitTrans: Call UXEditorContextChanged(): ArchiveDraft = True: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modSwift", "ArchiveDraft", "Execute", Err.number, Err.description, Err.source, "Unknown")
    SwiftRollback w, swiftPrimary: Err.Raise CLng(swiftPrimary("Number")), "ArchiveDraft", CStr(swiftPrimary("Description"))
End Function

Private Function RevisionValue(ByVal revisionID As Long, ByVal key As String) As Variant
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset
    Set r = ownedDB.OpenRecordset("SELECT ValueText, IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='" & key & "'", dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 114, "modSwift.RevisionValue", "Revision snapshot is incomplete."
    If r!IsNullValue Then RevisionValue = Null Else RevisionValue = r!ValueText
    r.Close
End Function

Public Function RestoreDraft(ByVal reportID As Long, ByVal revisionID As Long) As Long
    If Nz(DLookup("ReportID", "ReportRevision", "RevisionID=" & revisionID), 0) <> reportID Then Err.Raise vbObjectError + 115, "modSwift.RestoreDraft", "Revision belongs to another report."
    If RevisionValuesDisposed(revisionID) Then Err.Raise vbObjectError + 116, "modSwift.RestoreDraft", "This draft's contents were disposed of after signing. Only signed and current revisions can be restored."
    RestoreDraft = CommitDraft(reportID, RevisionValue(revisionID, "FormKind"), Nz(RevisionValue(revisionID, "Title"), ""), Nz(RevisionValue(revisionID, "MemberName"), ""), RevisionValue(revisionID, "PeriodStart"), RevisionValue(revisionID, "PeriodEnd"), Nz(RevisionValue(revisionID, "DraftNotes"), ""), "Restore revision " & revisionID, NarrativeRevisionValue(revisionID), AdministrationFromRevision(revisionID), TraitsFromRevision(revisionID), AssessmentFromRevision(revisionID), PromotionStatusFromRevision(revisionID), ReportingSeniorFromRevision(revisionID), ReportContextFromRevision(revisionID), MemberDetailsFromRevision(revisionID), ServiceDetailsFromRevision(revisionID), DecisionDetailsFromRevision(revisionID), ReportPeopleFromRevision(revisionID), PartyNamesFromRevision(revisionID), Block28FromRevision(revisionID), RevisionFields(revisionID))
End Function

Public Function DuplicateDraft(ByVal reportID As Long) As Long
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset, link As DAO.Recordset, failure As Long, source As String, message As String
    Set r = ownedDB.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 116, "modSwift.DuplicateDraft", "Report does not exist."
    'The copy is the same Sailor: keep the original's member link and its trust
    'state instead of matching the name again.
    Set link = ownedDB.OpenRecordset("SELECT m.MemberID,m.MatchState FROM MemberReport AS m INNER JOIN CommandMember AS c ON m.MemberID=c.MemberID WHERE c.IsArchived=False AND m.ReportID=" & reportID, dbOpenSnapshot)
    If Not link.EOF Then DirectoryCreationContext link!MemberID, Nz(link!MatchState, "Provisional"), "Duplicated report"
    link.Close
    On Error GoTo Failed
    DuplicateDraft = CommitDraft(0, r!FormKind, Left$("Copy of " & Nz(r!Title, ""), 160), Nz(r!MemberName, ""), r!PeriodStart, r!PeriodEnd, Nz(r!DraftNotes, ""), "Duplicate report " & reportID, r!OfficialComments, Array(r!OfficialName.Value, r!RankRate.Value, r!Designator.Value, r!UnitUIC.Value, r!ShipStation.Value), Array(r!Trait33.Value, r!Trait34.Value, r!Trait35.Value, r!Trait36.Value, r!Trait37.Value, r!Trait38.Value, r!Trait39.Value), Array(r!ObservationStatus.Value, r!PromotionRecommendation.Value), r!PromotionStatus.Value, Array(r!ReportingSeniorName.Value, r!ReportingSeniorGrade.Value, r!ReportingSeniorDesignator.Value, r!ReportingSeniorTitle.Value, r!ReportingSeniorUIC.Value), ReportContextFromRecordset(r), MemberDetailsFromRecordset(r), ServiceDetailsFromRecordset(r), DecisionDetailsFromRecordset(r), ReportPeopleFromRecordset(r), PartyNamesFromRecordset(r), StoredBlock28(reportID), SavedFields(reportID))
    DirectoryCreationContext 0
    r.Close
    Exit Function
Failed:
    failure = Err.Number: source = Err.Source: message = Err.Description
    DirectoryCreationContext 0
    Err.Raise failure, source, message
End Function

Public Function Startup() As Boolean
    Startup = SwiftAppStartup()
End Function

Public Function SwiftAppStartup() As Boolean
    On Error GoTo Failed
    If Nz(DLookup("SchemaVersion", "AppMetadata"), 0) <> SCHEMA_VERSION Then
        MsgBox "Unsupported Swift schema. Open this file with its matching application version.", vbCritical
        Application.Quit acQuitSaveNone
        Exit Function
    End If
    If CurrentProject.AllForms("frmEditor").IsLoaded Or UXGridPendingCount() > 0 Then Err.Raise 5, "modSwift.SwiftAppStartup", "Finish pending report edits before restarting contributor selection."
    If CurrentProject.AllForms("frmHome").IsLoaded Then DoCmd.Close acForm, "frmHome", acSaveNo
    ResetActiveContributor
    ShowContributorStartup
    UXLoadContext
    UXHideAccessObjects
    SwiftAppStartup = True: Exit Function
Failed:
    MsgBox "Swift NAVFIT could not start. Error " & Err.Number & ": " & Err.Description, vbExclamation, "Swift NAVFIT startup"
End Function

Public Function OpenScreen(ByVal formName As String) As Boolean
    On Error GoTo Failed
    If formName = "frmHelp" Then formName = "frmUXHelp"
    If Not HasActiveContributor() Then
        OpenScreen = ShowContributorStartup(): Exit Function
    End If
    If UXLegacyTool(formName) <> "" Then OpenScreen = UXLegacyToolOpen(UXLegacyTool(formName)): Exit Function
    If UXIsWorkspace(formName) Or formName = "frmBulkFolders" Or formName = "frmUXHelp" Then OpenScreen = UXNavigate(formName): Exit Function
    If formName = "frmSnippets" Then UXResetSnippetOrigin
    DoCmd.OpenForm formName
    Forms(formName).SetFocus
    OpenScreen = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation
End Function

Public Function SaveSetup() As Boolean
    On Error GoTo Failed
    ConfigureWorkspace Nz(Forms!frmSetup!txtCommand, ""), Nz(Forms!frmSetup!txtProfile, "")
    DoCmd.Close acForm, "frmSetup": DoCmd.OpenForm "frmHome": SaveSetup = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation
End Function

Public Function LoadHome() As Boolean
    Forms!frmHome!uxContributor.RowSource = "SELECT UserProfileID,DisplayName FROM UserProfile ORDER BY DisplayName"
    Forms!frmHome!uxContributor = ProfileID()
    Forms!frmHome!uxContributorStatus.Caption = ""
    Forms!frmHome!lblCommand.Caption = Nz(DLookup("CommandName", "Workspace"), "Swift NAVFIT")
    Forms!frmHome!lblProfile.Caption = "Contributor: " & Nz(DLookup("DisplayName", "UserProfile", "UserProfileID=" & ProfileID()), "")
    Forms!frmHome!cmdRecovery.Caption = "Recover pending edits (" & DCount("*", "RecoveryDraft", "AuthorProfileID=" & ProfileID()) & ")"
    UXDashboardRefresh
    LoadHome = True
End Function

Public Function LoadLibrary() As Boolean
    RefreshFolders
    If IsNull(Forms!frmLibrary!cboArchive) Then Forms!frmLibrary!cboArchive = "Active"
    UXMemberRankChoices Forms!frmLibrary, "Report"
    Call RefreshLibrary()
    LoadLibrary = True
End Function

Public Function EditSelected(Optional ByVal createNew As Boolean = False) As Boolean
    Dim id As Long
    If Not createNew Then
        id = Nz(Forms!frmLibrary!lstReports.Value, 0)
        If id = 0 Then MsgBox "Select a report first.", vbInformation: Exit Function
    End If
    If CurrentProject.AllForms("frmEditor").IsLoaded Then
        MsgBox "Close the current editor before opening another report.", vbInformation: Exit Function
    End If
    Call UXCaptureReportList
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    EditSelected = True
End Function

Public Function LoadEditor() As Boolean
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim f As Form, r As DAO.Recordset, id As Long, block As Long, seniorKey As Variant
    Call UXAssertView(UX_EDITOR_VIEW)
    UXFixedForgetLayout
    Set f = Forms!frmEditor: id = Val(Nz(f.OpenArgs, "0")): f!txtReportID = id
    f!cboKind = "EVAL"
    If id > 0 Then
        AssertReportCache id
        Set r = ownedDB.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
        f!cboKind = r!FormKind: f!txtTitle = r!Title: f!txtMember = r!MemberName
        f!txtStart = r!PeriodStart: f!txtEnd = r!PeriodEnd: f!txtNotes = r!DraftNotes: f!txtOfficialComments = r!OfficialComments
        f!txtOfficialName = r!OfficialName: f!txtRankRate = r!RankRate: f!txtDesignator = r!Designator: f!txtUnitUIC = r!UnitUIC: f!txtShipStation = r!ShipStation
        For block = 33 To 39: f.Controls("cboTrait" & block) = r.Fields("Trait" & block).Value: Next
        For Each seniorKey In ReportingSeniorKeys(): f.Controls("txt" & seniorKey) = r.Fields(seniorKey).Value: Next
        For Each seniorKey In ReportContextKeys(): f.Controls("txt" & seniorKey) = r.Fields(seniorKey).Value: Next
        For Each seniorKey In MemberDetailsKeys(): f.Controls("txt" & seniorKey) = r.Fields(seniorKey).Value: Next
        For Each seniorKey In ServiceDetailsKeys(): f.Controls("txt" & seniorKey) = r.Fields(seniorKey).Value: Next
        For Each seniorKey In DecisionDetailsKeys(): f.Controls("txt" & seniorKey) = r.Fields(seniorKey).Value: Next
        For Each seniorKey In ReportPeopleKeys(): f.Controls("txt" & seniorKey) = r.Fields(seniorKey).Value: Next
        For Each seniorKey In PartyNamesKeys(): f.Controls("txt" & seniorKey) = r.Fields(seniorKey).Value: Next
        f!txtCommandEmployment = StoredBlock28(id)
        Dim fieldKeys As Variant, fieldIndex As Long
        fieldKeys = DraftFieldKeys()
        For fieldIndex = 79 To UBound(fieldKeys)
            f.Controls(UXFieldControl(CStr(fieldKeys(fieldIndex)))) = r.Fields(CStr(fieldKeys(fieldIndex))).Value
        Next
        f!cboPromotionStatus = r!PromotionStatus
        f!cboObservationStatus = r!ObservationStatus: f!cboPromotionRecommendation = r!PromotionRecommendation
        f!lblSaved.Caption = "Saved revision " & Nz(DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & r!CurrentRevisionID), "?") & " at " & r!ModifiedAt
        f!cmdSave.Enabled = Not r!IsArchived
        f!cmdRestore.Enabled = Not r!IsArchived
        r.Close: f!cboKind.Enabled = False
        Select Case NAVFITImportScope(id)
            Case "Full report": f!lblImportOrigin.Caption = "Full NAVFIT report import: supported official fields mapped; native-only values retained in provenance."
            Case "Narrative only": f!lblImportOrigin.Caption = "Imported narrative only. Other official fields remain unmapped; source values retained locally."
        End Select
    End If
    If id = 0 Then UXSeedGlobalDefaults
    Call InitializeTraitEditor()
    Call RefreshNarrativeLabel()
    Call RefreshRevisions
    Call BeginEditorRecovery()
    Call UXEditorStart()
    LoadEditor = True
End Function

Private Sub RefreshRevisions()
    Dim f As Form
    Set f = Forms!frmEditor
    f!lstRevisions.RowSource = "SELECT RevisionID, RevisionNumber AS [Revision], CommittedAt AS [Saved], IIf(Nz(ValuesDisposed,False),'[Draft contents disposed] ','') & Reason AS [Note] FROM ReportRevision WHERE ReportID=" & Nz(f!txtReportID, 0) & " ORDER BY RevisionNumber DESC"
    f!lstRevisions.Requery
End Sub

Public Function SaveEditor() As Boolean
    Dim f As Form, id As Long, w As DAO.Workspace, trans As Boolean, committed As Boolean, stage As String
    On Error GoTo Failed
    Set f = Forms!frmEditor
    stage = "Capture and validate pending values"
    UXFlushSurface
    UXFormatTypedFields
    Call GuardEditorBaseline()
    Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    stage = "Write draft revision"
    id = CommitDraft(Nz(f!txtReportID, 0), Nz(f!cboKind, ""), Nz(f!txtTitle, ""), Nz(f!txtMember, ""), f!txtStart, f!txtEnd, Nz(f!txtNotes, ""), "Draft save", f!txtOfficialComments, AdministrationFromEditor(), TraitsFromEditor(), AssessmentFromEditor(), f!cboPromotionStatus.Value, ReportingSeniorFromEditor(), ReportContextFromEditor(), MemberDetailsFromEditor(), ServiceDetailsFromEditor(), DecisionDetailsFromEditor(), ReportPeopleFromEditor(), PartyNamesFromEditor(), f!txtCommandEmployment.Value, EditorPendingFields())
    stage = "Clear committed checkpoint"
    Call DiscardEditorCheckpoint()
    stage = "Commit save"
    w.CommitTrans: trans = False: committed = True
    stage = "Refresh saved report display"
    f!txtReportID = id: f!cboKind.Enabled = False: f!lblSaved.Caption = "Saved revision " & Nz(DMax("RevisionNumber", "ReportRevision", "ReportID=" & id), "?") & " at " & Format$(Now(), "h:nn:ss AM/PM") & " - local draft"
    RefreshRevisions
    If CurrentProject.AllForms("frmLibrary").IsLoaded Then LoadLibrary
    Call ResetEditorBaseline()
    Call RefreshTraitSummary()
    f!lblRecovery.Caption = "Recovery: no unsaved edits to checkpoint."
    If UXHasControl(f, "uxReviewSetup") Then UXHeaderReviewState
    SaveEditor = True: Exit Function
Failed:
    Dim swiftPrimary As Object
    Set swiftPrimary = SwiftFailure("modSwift", "SaveEditor", stage, Err.Number, Err.Description, Err.Source, IIf(committed, "Committed", IIf(trans, "Active", "Not started")))
    Dim message As String
    message = CStr(swiftPrimary("Description"))
    If trans Then SwiftRollback w, swiftPrimary
    On Error GoTo 0
    message = SwiftFailureText(swiftPrimary)
    If Not UXShowSaveError(message, committed) Then MsgBox message, vbExclamation
    SaveEditor = committed
End Function

Public Function CloseEditor() As Boolean
    CloseEditor = UXLeaveEditor()
End Function

Public Function RestoreSelected() As Boolean
    Dim id As Long, revisionID As Long
    On Error GoTo Failed
    id = Nz(Forms!frmEditor!txtReportID, 0): revisionID = Nz(Forms!frmEditor!lstRevisions, 0)
    If revisionID = 0 Then MsgBox "Select a revision first.": Exit Function
    If MsgBox("Restore selected values as a new revision? Unsaved entries will be discarded.", vbOKCancel + vbQuestion) <> vbOK Then Exit Function
    RestoreDraft id, revisionID
    Call DiscardEditorCheckpoint()
    DoCmd.Close acForm, "frmEditor", acSaveNo
    DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
    RestoreSelected = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation
End Function

Public Function LibraryAction(ByVal action As String) As Boolean
    Dim id As Long
    On Error GoTo Failed
    id = Nz(Forms!frmLibrary!lstReports, 0)
    If id = 0 Then MsgBox "Select a report first.": Exit Function
    If CurrentProject.AllForms("frmEditor").IsLoaded Then MsgBox "Close the editor before changing library reports.": Exit Function
    Select Case action
        Case "Archive": ArchiveDraft id, True
        Case "Recover": ArchiveDraft id, False
        Case "Duplicate": DuplicateDraft id
    End Select
    Call LoadLibrary()
    LibraryAction = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation
End Function

Public Function LoadSupport() As Boolean
    Forms!frmSupport!txtDiagnostics = "Swift " & APP_VERSION & vbCrLf & "Schema " & SCHEMA_VERSION & vbCrLf & "Access " & Application.Version & " build " & Application.Build & vbCrLf & "File: " & CurrentProject.FullName & vbCrLf & WorkspaceSizeText(WorkspaceFileBytes()) & vbCrLf & "NAVFIT adapter: 75 mapped saved values; 128 native columns retained; report/group export and guarded return preview; verify files in NAVFIT98A" & vbCrLf & "Policy validation: limited live report/group diagnostics with source references; not Navy-ready certification" & vbCrLf & "PDF import: limited structured-field mapping with preview and provenance" & vbCrLf & "Timestamps: workstation local time; offset capture not yet implemented" & vbCrLf & "Local profiles and audit records are not strong authentication or tamper-resistant security."
    LoadSupport = True
End Function



Private Function NarrativeRevisionValue(ByVal revisionID As Long) As Variant
    'Schema 1/2 revisions predate the official narrative field.
    If DCount("*", "RevisionValue", "RevisionID=" & revisionID & " AND FieldKey='OfficialComments'") = 0 Then
        NarrativeRevisionValue = Null
    Else
        NarrativeRevisionValue = RevisionValue(revisionID, "OfficialComments")
    End If
End Function

Private Function StoredOfficialNarrative(ByVal reportID As Long) As Variant
    Dim ownedDB As DAO.Database
    Set ownedDB = CurrentDb
    Dim r As DAO.Recordset
    Set r = ownedDB.OpenRecordset("SELECT OfficialComments FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 109, "modSwift.StoredOfficialNarrative", "Report does not exist."
    StoredOfficialNarrative = r!OfficialComments
    r.Close
End Function

Public Function RefreshNarrativeLabel() As Boolean
    Dim block As String
    Select Case Nz(Forms!frmEditor!cboKind, "")
        Case "CHIEFEVAL": block = "40"
        Case "EVAL": block = "43"
        Case "FITREP": block = "41"
        Case Else: block = "?"
    End Select
    Forms!frmEditor!lblOfficialComments.Caption = "Official narrative - block " & block & " (separate from draft notes)"
    Forms!frmEditor!lblNarrativeCount.Caption = CStr(Len(Nz(Forms!frmEditor!txtOfficialComments, ""))) & " characters"
    RefreshNarrativeLabel = True
End Function

Public Function Block28FromRevision(ByVal revisionID As Long) As Variant
    If Not MajorRevisionHasField(revisionID, "CommandEmployment") Then
        If RevisionSnapshotSchema(revisionID) >= 24 Then Err.Raise vbObjectError + 119, "modSwift.Block28FromRevision", "Block 28 is missing from this revision."
        Block28FromRevision = Null
    Else
        Block28FromRevision = MajorRevisionValue(revisionID, "CommandEmployment")
    End If
End Function

Public Function StoredBlock28(ByVal reportID As Long) As Variant
    Dim db As DAO.Database, r As DAO.Recordset, receipt As Variant, nativeID As Long
    StoredBlock28 = Null: If reportID = 0 Then Exit Function
    Set db = CurrentDb
    StoredBlock28 = DLookup("CommandEmployment", "Report", "ReportID=" & reportID)
    If Not IsNull(StoredBlock28) Then Exit Function
    Set r = db.OpenRecordset("SELECT c.RevisionChangeID FROM RevisionChange AS c INNER JOIN ReportRevision AS v ON c.RevisionID=v.RevisionID WHERE v.ReportID=" & reportID & " AND c.FieldKey='CommandEmployment'", dbOpenSnapshot)
    If Not r.EOF Then r.Close: Exit Function
    r.Close
    'Only reports whose first revision predates Block 28 use the native value. Since
    'L03 a first save no longer logs an empty Block 28, so check its snapshot too.
    Set r = db.OpenRecordset("SELECT TOP 1 v.FieldKey FROM RevisionValue AS v WHERE v.FieldKey='CommandEmployment' AND v.RevisionID=(SELECT Min(RevisionID) FROM ReportRevision WHERE ReportID=" & reportID & ")", dbOpenSnapshot)
    If Not r.EOF Then r.Close: Exit Function
    r.Close
    receipt = MajorScalar("SELECT Max(ImportReceiptID) FROM ImportReceipt WHERE ReportID=" & reportID & " AND IsUndone=False AND AdapterVersion='navfit-full-v1/local-CStr'")
    If IsNull(receipt) Then Exit Function
    nativeID = Nz(DLookup("ExternalReportID", "ImportReceipt", "ImportReceiptID=" & CLng(receipt)), 0)
    Set r = db.OpenRecordset("SELECT RawValue,IsNullValue FROM ImportRawValue WHERE ImportReceiptID=" & CLng(receipt) & " AND SourceRow=" & nativeID & " AND SourceTable IN ('Reports','RawReports') AND FieldKey='Achievements' AND DAOType=12", dbOpenSnapshot)
    If Not r.EOF Then
        If Not r!IsNullValue Then StoredBlock28 = Nz(r!RawValue.Value, "")
    End If
    r.Close
End Function
