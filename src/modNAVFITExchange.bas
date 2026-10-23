Attribute VB_Name = "modNAVFITExchange"
Option Compare Binary
Option Explicit
Private Type ReturnedPreviewContext
    Path As String
    Hash As String
    Baseline As String
    BaselineHash As String
    Report As Long
    Revision As Long
    External As Long
    Century As Long
    Values As Object
    Decisions As Object
    Before As Object
    Frozen As Object
    NativeText As String
    NativeBase As Object
    NativeIncoming As Object
    StageGuid As String
End Type
Private returnContext As ReturnedPreviewContext















Private reportReadiness As Collection
Private groupReadiness As Collection
Public Function NAVFITLastGroupReadiness() As Collection
    Set NAVFITLastGroupReadiness = groupReadiness
End Function
' Return the findings from the saved revision actually exported by the service.
Public Function NAVFITLastReportReadiness() As Collection
    Set NAVFITLastReportReadiness = reportReadiness
End Function
Public Function NAVFITGroupPreview(ByVal groupID As Long, Optional ByVal usePSG As Boolean = False) As Object
    Dim result As Object, members As DAO.Recordset, values As Object, first As Object, key As Variant, id As Long, code As Variant, counts(0 To 5) As Long, psg As Object, text As String, token As String, title As String
    Dim readiness As Collection, savedOnly As Object
    If Nz(MajorScalar("SELECT IsArchived FROM SummaryGroup WHERE GroupID=" & groupID), True) Then Err.Raise vbObjectError + 5350, , "Choose an active summary group."
    title = CStr(MajorScalar("SELECT Title FROM SummaryGroup WHERE GroupID=" & groupID))
    If Len(title) > 50 Then Err.Raise vbObjectError + 5350, , "NAVFIT folder names allow 50 characters. Shorten this group name before exporting."
    Set result = CreateObject("Scripting.Dictionary"): Set psg = RSGroupPSG(groupID)
    Set members = CurrentDb.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & groupID & " ORDER BY ReportID", dbOpenSnapshot)
    If members.EOF Then members.Close: Err.Raise vbObjectError + 5350, , "Group has no reports."
    text = title & vbCrLf & "Saved members:" & vbCrLf
    Do While Not members.EOF
        id = members!ReportID: NAVFITCheckTarget id, False: Set values = SavedFields(id)
        If values("FormKind") <> MajorScalar("SELECT FormKind FROM SummaryGroup WHERE GroupID=" & groupID) Then Err.Raise vbObjectError + 5350, , "A member's report type differs from the group."
        If first Is Nothing Then
            Set first = values
            If Len(Nz(values("ReportingSeniorSSN"), "")) <> 9 Or IsNull(values("PeriodEnd")) Then Err.Raise vbObjectError + 5350, , "Group export requires a reporting senior SSN and report end date."
        Else
            For Each key In Array("ReportingSeniorSSN", "ReportingSeniorName", "ReportingSeniorGrade", "ReportingSeniorDesignator", "ReportingSeniorTitle", "ReportingSeniorUIC", "PeriodEnd", "ReportRegular", "ReportConcurrent", "ReportOperationalCommander", "ComponentActive", "ComponentTAR", "ComponentInactive", "ComponentATADSW")
                If Not SameFieldValue(first(CStr(key)), values(CStr(key))) Then Err.Raise vbObjectError + 5350, , "Group members differ in " & UXFieldCaption(CStr(key), CStr(values("FormKind"))) & ". Review membership or shared blocks before exporting."
            Next
        End If
        code = NAVFITEncode(values("PromotionRecommendation"), "promotion")
        If IsNull(code) Then Err.Raise vbObjectError + 5350, , "Choose a promotion recommendation (including NOB) for " & UXMemberLabel(id) & "."
        counts(CLng(code)) = counts(CLng(code)) + 1
        text = text & UXMemberLabel(id) & " | " & values("PromotionRecommendation") & vbCrLf
        members.MoveNext
    Loop
    members.Close
    If DCount("*", "GroupMember", "GroupID=" & groupID) > 999 Then Err.Raise vbObjectError + 5350, , "Native summary-count fields allow at most 999 reports."
    text = text & vbCrLf & "Counts: SP " & counts(1) & "; Progressing " & counts(2) & "; Promotable " & counts(3) & "; MP " & counts(4) & "; EP " & counts(5) & "; NOB " & counts(0) & vbCrLf
    If usePSG Then
        If IsNull(psg("Value")) Or psg("Estimated") Or psg("Partial") Then Err.Raise vbObjectError + 5350, , "Exporting PSG RSCA requires a complete group and an exact historical trait sum. " & psg("Reason")
        text = text & "Write PSG RSCA " & Format$(psg("Value"), "0.00") & " to each report. " & psg("Reason") & vbCrLf
    Else
        text = text & "Preserve each imported report's native RSCA; new drafts have no native RSCA. PSG RSCA is not inserted." & vbCrLf
    End If
    text = text & "One native folder; saved revisions only. Native validation is required. Group-return reconciliation is not yet supported; keep this frozen baseline."
    token = GroupStamp(groupID) & "|" & CStr(usePSG) & "|" & NAVFITDisplay(psg("Value")) & "|" & psg("Reason") & "|" & CStr(psg("Estimated")) & "|" & CStr(psg("Partial"))
    Set readiness = ValidateLiveGroup(groupID, , , False)
    Set groupReadiness = readiness
    text = text & vbCrLf & vbCrLf & ReadinessSummary(readiness)
    result.Add "Readiness", readiness
    result.Add "Token", token: result.Add "Text", text: result.Add "Title", title: result.Add "PSG", psg("Value")
    For id = 0 To 5: result.Add "Count" & id, counts(id): Next
    Set NAVFITGroupPreview = result
End Function
Public Function NAVFITExportGroup(ByVal groupID As Long, ByVal destination As String, ByVal blankTemplate As String, ByVal expectedToken As String, Optional ByVal usePSG As Boolean = False) As String
    Dim preview As Object, native As DAO.Database, source As DAO.Database, baseline As DAO.Database, members As DAO.Recordset, report As DAO.Recordset, meta As DAO.Recordset, fields As DAO.Recordset, f As DAO.Field, values As Object, key As Variant, value As Variant
    Dim pending As String, manifest As String, pendingManifest As String, staging As String, operation As Long, id As Long, external As Long, receipt As Variant, names As String, selection As String, newID As Long, century As Long, index As Long, countFields As Variant, failure As Long, message As String
    On Error GoTo Failed
    Set preview = NAVFITGroupPreview(groupID, usePSG)
    If expectedToken <> preview("Token") Then Err.Raise vbObjectError + 5351, , "Group, saved reports or RSCA changed. Preview again."
    If LCase$(Right$(destination, 6)) <> ".accdb" Then Err.Raise vbObjectError + 5351, , "Export filename must end in .accdb."
    manifest = destination & ".swift-baseline.accdb"
    If Len(Dir$(destination)) > 0 Or Len(Dir$(manifest)) > 0 Then Err.Raise vbObjectError + 5351, , "Choose a new export filename; existing files are not overwritten."
    operation = BeginOperation("NAVFIT group export", destination)
    pending = destination & "." & Replace(Replace(NewIdentity(), "{", ""), "}", "") & ".pending.accdb": pendingManifest = pending & ".baseline.accdb"
    FileCopy blankTemplate, pending: Set native = DBEngine.OpenDatabase(pending): ValidateNAVFITDatabase native
    native.Execute "DELETE FROM Reports WHERE ReportType Is Not Null", dbFailOnError
    native.Execute "DELETE FROM Folders WHERE FolderID<>1", dbFailOnError
    Set report = native.OpenRecordset("SELECT * FROM Folders WHERE FolderID=1", dbOpenDynaset)
    If report.EOF Then Err.Raise vbObjectError + 5351, , "Blank template has no root folder."
    report.Edit: report!FolderName = preview("Title"): report!Parent = 0: report!AutoSummary = True
    countFields = Array("SummarySP", "SummaryProg", "SummaryProm", "SummaryMP", "SummaryEP")
    For index = 1 To 5: report.Fields(countFields(index - 1)) = CStr(preview("Count" & index)): Next
    report.Update: report.Close: Set report = Nothing
    Set members = CurrentDb.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & groupID & " ORDER BY ReportID", dbOpenSnapshot)
    ' Native identities are deterministic within this package, not guessed from SSNs.
    newID = 1
    Do While Not members.EOF
        id = members!ReportID: Set values = SavedFields(id): newID = newID + 1
        receipt = MajorScalar("SELECT Max(ImportReceiptID) FROM ImportReceipt WHERE ReportID=" & id & " AND IsUndone=False AND AdapterVersion='navfit-full-v1/local-CStr'")
        If Not IsNull(receipt) Then
            staging = pending & ".member-" & id & ".accdb": NAVFITRebuildSource CLng(receipt), blankTemplate, staging
            external = CLng(MajorScalar("SELECT ExternalReportID FROM ImportReceipt WHERE ImportReceiptID=" & CLng(receipt)))
            names = "": selection = ""
            For Each f In native.TableDefs("Reports").Fields
                If Len(names) > 0 Then names = names & ",": selection = selection & ","
                names = names & "[" & f.Name & "]"
                Select Case f.Name
                    Case "ReportID": selection = selection & CStr(newID)
                    Case "Parent": selection = selection & "'a 1'"
                    Case Else: selection = selection & "[" & f.Name & "]"
                End Select
            Next
            native.Execute "INSERT INTO Reports (" & names & ") SELECT " & selection & " FROM Reports IN " & QuoteSQL(staging) & " WHERE ReportID=" & external, dbFailOnError
            Kill staging: staging = ""
            Set report = native.OpenRecordset("SELECT * FROM Reports WHERE ReportID=" & newID, dbOpenDynaset)
            If report.EOF Then Err.Raise vbObjectError + 5351, , "Preserved native report is missing."
            report.Edit
        Else
            native.Execute "INSERT INTO Reports (ReportID,Parent) VALUES (" & newID & ",'a 1')", dbFailOnError
            Set report = native.OpenRecordset("SELECT * FROM Reports WHERE ReportID=" & newID, dbOpenDynaset): report.Edit
        End If
        NAVFITWriteFields report, values
        countFields = Array("SummarySP", "SummaryProg", "SummaryProm", "SummaryMP", "SummaryEP")
        For index = 1 To 5: report.Fields(countFields(index - 1)) = CStr(preview("Count" & index)): Next
        If usePSG Then report!RSCA = CDec(Round(CDbl(preview("PSG")), 2))
        report.Update: report.Close: Set report = Nothing: members.MoveNext
    Loop
    members.Close: Set members = Nothing
    VerifyNativeGroupSummary native, preview
    native.Close: Set native = Nothing
    FileCopy pending, pendingManifest: Set baseline = DBEngine.OpenDatabase(pendingManifest)
    baseline.Execute "CREATE TABLE SwiftGroupExport (WorkspaceGuid TEXT(38), GroupGuid TEXT(38), GroupID LONG, GroupToken LONGTEXT, OperationID LONG, UsePSG BIT)", dbFailOnError
    baseline.Execute "CREATE TABLE SwiftGroupReport (SwiftReportID LONG PRIMARY KEY, ReportGuid TEXT(38), BaseRevisionID LONG, ExternalReportID LONG, SourceReceiptID LONG, CounselingCentury LONG)", dbFailOnError
    baseline.Execute "CREATE TABLE SwiftGroupValue (SwiftReportID LONG, FieldKey TEXT(64), ValueText LONGTEXT, IsNullValue BIT, IsDateValue BIT, CONSTRAINT PK_SwiftGroupValue PRIMARY KEY (SwiftReportID,FieldKey))", dbFailOnError
    Set meta = baseline.OpenRecordset("SwiftGroupExport", dbOpenDynaset): meta.AddNew
    meta!WorkspaceGuid = MajorScalar("SELECT WorkspaceGuid FROM Workspace WHERE WorkspaceID=1"): meta!GroupGuid = MajorScalar("SELECT GroupGuid FROM SummaryGroup WHERE GroupID=" & groupID)
    meta!GroupID = groupID: meta!GroupToken = expectedToken: meta!OperationID = operation: meta!UsePSG = usePSG: meta.Update: meta.Close: Set meta = Nothing
    Set members = CurrentDb.OpenRecordset("SELECT ReportID FROM GroupMember WHERE GroupID=" & groupID & " ORDER BY ReportID", dbOpenSnapshot)
    Set meta = baseline.OpenRecordset("SwiftGroupReport", dbOpenDynaset): Set fields = baseline.OpenRecordset("SwiftGroupValue", dbOpenDynaset): newID = 1
    Do While Not members.EOF
        id = members!ReportID: newID = newID + 1: Set values = SavedFields(id): century = 0
        If Len(Nz(values("CounselingDate"), "")) = 10 Then century = (CLng(Left$(values("CounselingDate"), 4)) \ 100) * 100
        meta.AddNew: meta!SwiftReportID = id: meta!ReportGuid = MajorScalar("SELECT ReportGuid FROM Report WHERE ReportID=" & id): meta!BaseRevisionID = CurrentRevision(id): meta!ExternalReportID = newID: meta!CounselingCentury = century
        meta!SourceReceiptID = MajorScalar("SELECT Max(ImportReceiptID) FROM ImportReceipt WHERE ReportID=" & id & " AND IsUndone=False AND AdapterVersion='navfit-full-v1/local-CStr'"): meta.Update
        For Each key In values.Keys
            value = values(key): fields.AddNew: fields!SwiftReportID = id: fields!FieldKey = key: fields!IsNullValue = IsNull(value): fields!IsDateValue = (VarType(value) = vbDate)
            If Not IsNull(value) Then
                If VarType(value) = vbDate Then fields!ValueText = CStr(CDbl(value)) Else fields!ValueText = CStr(value)
            End If
            fields.Update
        Next
        members.MoveNext
    Loop
    members.Close: Set members = Nothing: meta.Close: Set meta = Nothing: fields.Close: Set fields = Nothing: baseline.Close: Set baseline = Nothing
    Set preview = NAVFITGroupPreview(groupID, usePSG)
    If expectedToken <> preview("Token") Then Err.Raise vbObjectError + 5351, , "Group changed during export. Pending files have not been published."
    FinishOperation operation, "Prepared", "Pending group export: " & pending & "; baseline: " & pendingManifest
    Name pendingManifest As manifest: Name pending As destination
    FinishOperation operation, "Completed", "Baseline SHA256 " & SourceFileHash(manifest)
    NAVFITExportGroup = destination: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not report Is Nothing Then report.Close
    If Not members Is Nothing Then members.Close
    If Not meta Is Nothing Then meta.Close
    If Not fields Is Nothing Then fields.Close
    If Not native Is Nothing Then native.Close
    If Not baseline Is Nothing Then baseline.Close
    If Len(staging) > 0 Then Kill staging
    If operation > 0 Then FinishOperation operation, "Failed", message & "; pending: " & pending
    On Error GoTo 0
    Err.Raise failure, "NAVFITExportGroup", message
End Function
Public Function NAVFITExportReport(ByVal reportID As Long, ByVal destination As String, ByVal blankTemplate As String) As String
    Dim db As DAO.Database, native As DAO.Database, baseline As DAO.Database, r As DAO.Recordset, meta As DAO.Recordset, fields As DAO.Recordset, values As Object
    Dim receipt As Variant, source As String, external As Long, revision As Long, operation As Long, pending As String, manifest As String, pendingManifest As String, key As Variant, v As Variant, century As Long
    Dim failure As Long, message As String
    On Error GoTo Failed
    NAVFITCheckTarget reportID, False
    Set values = SavedFields(reportID): revision = CurrentRevision(reportID)
    Set reportReadiness = ValidateLiveFields(values, reportID)
    If Len(Dir$(destination)) > 0 Or Len(Dir$(destination & ".swift-baseline.accdb")) > 0 Then Err.Raise vbObjectError + 5300, , "Choose a new export filename; existing files are not overwritten."
    If LCase$(Right$(destination, 6)) <> ".accdb" Then Err.Raise vbObjectError + 5300, , "Export filename must end in .accdb."
    receipt = MajorScalar("SELECT Max(ImportReceiptID) FROM ImportReceipt WHERE ReportID=" & reportID & " AND IsUndone=False AND AdapterVersion='navfit-full-v1/local-CStr'")
    source = blankTemplate
    If Not IsNull(receipt) Then
        source = Nz(MajorScalar("SELECT SourcePath FROM ImportReceipt WHERE ImportReceiptID=" & CLng(receipt)), "")
        external = CLng(MajorScalar("SELECT ExternalReportID FROM ImportReceipt WHERE ImportReceiptID=" & CLng(receipt)))
    End If
    If StrComp(source, destination, vbTextCompare) = 0 Or StrComp(CurrentProject.FullName, destination, vbTextCompare) = 0 Then Err.Raise vbObjectError + 5300, , "Choose a distinct export destination."
    operation = BeginOperation("NAVFIT export", destination)
    pending = destination & "." & Replace(Replace(NewIdentity(), "{", ""), "}", "") & ".pending.accdb"
    manifest = destination & ".swift-baseline.accdb": pendingManifest = pending & ".baseline.accdb"
    If IsNull(receipt) Then
        FileCopy blankTemplate, pending
    Else
        NAVFITRebuildSource CLng(receipt), blankTemplate, pending
    End If
    Set native = DBEngine.OpenDatabase(pending)
    ValidateNAVFITDatabase native
    If external > 0 Then
        native.Execute "DELETE FROM Reports WHERE ReportType Is Not Null AND ReportID<>" & external, dbFailOnError
        Set r = native.OpenRecordset("SELECT * FROM Reports WHERE ReportID=" & external, dbOpenDynaset)
        If r.EOF Then Err.Raise vbObjectError + 5302, , "Source identity is missing."
        If NAVFITKind(r!ReportType) <> values("FormKind") Then Err.Raise vbObjectError + 5302, , "Source identity has a different report type."
        r.Edit
    Else
        native.Execute "DELETE FROM Reports WHERE ReportType Is Not Null", dbFailOnError
        Set r = native.OpenRecordset("Reports", dbOpenDynaset): r.AddNew: r!Parent = "a 1"
    End If
    NAVFITWriteFields r, values: r.Update
    If external = 0 Then r.Bookmark = r.LastModified: external = r!ReportID
    r.Close: native.Close: Set native = Nothing
    FileCopy pending, pendingManifest
    Set baseline = DBEngine.OpenDatabase(pendingManifest)
    baseline.Execute "CREATE TABLE SwiftExport (WorkspaceGuid TEXT(38), ReportGuid TEXT(38), SwiftReportID LONG, BaseRevisionID LONG, ExternalReportID LONG, CounselingCentury LONG, OperationID LONG, SourceReceiptID LONG)", dbFailOnError
    baseline.Execute "CREATE TABLE SwiftValue (FieldKey TEXT(64) CONSTRAINT PK_SwiftValue PRIMARY KEY, ValueText LONGTEXT, IsNullValue BIT, IsDateValue BIT)", dbFailOnError
    Set meta = baseline.OpenRecordset("SwiftExport", dbOpenDynaset)
    meta.AddNew: meta!WorkspaceGuid = MajorScalar("SELECT WorkspaceGuid FROM Workspace WHERE WorkspaceID=1")
    meta!ReportGuid = MajorScalar("SELECT ReportGuid FROM Report WHERE ReportID=" & reportID)
    meta!SwiftReportID = reportID: meta!BaseRevisionID = revision: meta!ExternalReportID = external: meta!OperationID = operation
    century = 0
    If Len(Nz(values("CounselingDate"), "")) = 10 Then century = (CLng(Left$(values("CounselingDate"), 4)) \ 100) * 100
    meta!CounselingCentury = century: meta!SourceReceiptID = receipt: meta.Update: meta.Close
    Set fields = baseline.OpenRecordset("SwiftValue", dbOpenDynaset)
    For Each key In values.Keys
        v = values(key): fields.AddNew: fields!FieldKey = key: fields!IsNullValue = IsNull(v): fields!IsDateValue = (VarType(v) = vbDate)
        If Not IsNull(v) Then
            If VarType(v) = vbDate Then fields!ValueText = CStr(CDbl(v)) Else fields!ValueText = CStr(v)
        End If
        fields.Update
    Next
    fields.Close: baseline.Close: Set baseline = Nothing
    If CurrentRevision(reportID) <> revision Then Err.Raise vbObjectError + 5303, , "Saved report changed during export. Pending files have not been published."
    FinishOperation operation, "Prepared", "Pending export: " & pending & "; baseline: " & pendingManifest
    Name pendingManifest As manifest
    Name pending As destination
    FinishOperation operation, "Completed", "Baseline SHA256 " & SourceFileHash(manifest)
    NAVFITExportReport = destination: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not r Is Nothing Then r.Close
    If Not meta Is Nothing Then meta.Close
    If Not fields Is Nothing Then fields.Close
    If Not native Is Nothing Then native.Close
    If Not baseline Is Nothing Then baseline.Close
    If operation > 0 Then FinishOperation operation, "Failed", message & "; pending: " & pending & "; baseline: " & pendingManifest
    On Error GoTo 0
    Err.Raise failure, "NAVFITExportReport", message
End Function
Public Sub NAVFITCheckTarget(ByVal reportID As Long, Optional ByVal writing As Boolean = True)
    Dim r As DAO.Recordset, db As DAO.Database
    If CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 5304, , "Save and close the report editor first."
    If DCount("*", "RecoveryDraft", "ReportID=" & reportID) > 0 Then Err.Raise vbObjectError + 5304, , "Resolve pending recovery for this report first."
    Set db = CurrentDb: Set r = db.OpenRecordset("SELECT IsArchived,WorkflowState FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 5304, , "Report no longer exists."
    If r!IsArchived Or (writing And r!WorkflowState <> "Draft") Then r.Close: Err.Raise vbObjectError + 5304, , "Recover archived reports or reopen local approval before applying exchange changes."
    r.Close
End Sub
Public Function NAVFITPreviewReturned(ByVal path As String, ByVal baselinePath As String, Optional ByVal counselingCentury As Long = 0, Optional ByVal resumeToken As String = "", Optional ByVal restoredChoices As Object = Nothing) As Collection
    Dim native As DAO.Database, baseline As DAO.Database, meta As DAO.Recordset, fields As DAO.Recordset, source As DAO.Recordset, baseValues As Object, row As Variant, parts As Variant, key As String, result As New Collection, operation As Long, outcome As String, value As Variant, receipt As Variant, nativeField As DAO.Field, mappedNames As Object, frozenNative As DAO.Recordset
    Dim failure As Long, message As String, stage As String
    On Error GoTo Failed
    stage = "Reset live preview": NAVFITClearReturned False
    returnContext.Hash = SourceFileHash(path): returnContext.BaselineHash = SourceFileHash(baselinePath)
    stage = "Read sealed baseline": Set baseline = DBEngine.OpenDatabase(baselinePath, False, True)
    If NAVFITHasTable(baseline, "SwiftGroupExport") Then Err.Raise vbObjectError + 5310, , "This is a group export baseline. Group returns are not supported by the single-report return screen."
    Set meta = baseline.OpenRecordset("SwiftExport", dbOpenSnapshot)
    If meta.EOF Then Err.Raise vbObjectError + 5310, , "Choose the export's Swift baseline file."
    returnContext.Report = meta!SwiftReportID: returnContext.External = meta!ExternalReportID: returnContext.Century = meta!CounselingCentury
    If counselingCentury <> 0 Then
        If counselingCentury <> 1900 And counselingCentury <> 2000 Then Err.Raise vbObjectError + 5319, , "Choose 1900 or 2000 for the returned counseling year."
        returnContext.Century = counselingCentury
    End If
    operation = meta!OperationID
    If meta!WorkspaceGuid <> MajorScalar("SELECT WorkspaceGuid FROM Workspace WHERE WorkspaceID=1") Or meta!ReportGuid <> MajorScalar("SELECT ReportGuid FROM Report WHERE ReportID=" & returnContext.Report) Then Err.Raise vbObjectError + 5311, , "Export baseline belongs to another workspace or report."
    outcome = Nz(MajorScalar("SELECT Outcome FROM OperationJournal WHERE OperationID=" & operation & " AND OperationKind='NAVFIT export' AND OperationState='Completed'"), "")
    If outcome <> "Baseline SHA256 " & returnContext.BaselineHash Then Err.Raise vbObjectError + 5312, , "Baseline is modified or its export did not complete."
    receipt = MajorScalar("SELECT Max(ImportReceiptID) FROM ImportReceipt WHERE ReportID=" & returnContext.Report & " AND IsUndone=False AND AdapterVersion='navfit-full-v1/local-CStr'")
    If Not SameFieldValue(receipt, meta!SourceReceiptID) Then Err.Raise vbObjectError + 5318, , "A newer native import replaced the preserved source values. Export a fresh baseline before reconciling this return."
    stage = "Check current Swift target": NAVFITCheckTarget returnContext.Report: returnContext.Revision = CurrentRevision(returnContext.Report)
    Set returnContext.Before = SavedFields(returnContext.Report): Set baseValues = CreateObject("Scripting.Dictionary")
    Set fields = baseline.OpenRecordset("SwiftValue", dbOpenSnapshot)
    Do While Not fields.EOF
        value = Null
        If Not fields!IsNullValue Then
            value = Nz(fields!ValueText, "")
            If fields!IsDateValue Then value = CDate(CDbl(value))
        End If
        baseValues.Add CStr(fields!FieldKey), value: fields.MoveNext
    Loop
    fields.Close: Set fields = Nothing: meta.Close: Set meta = Nothing
    stage = "Check native file structure": Set native = DBEngine.OpenDatabase(path, False, True): ValidateNAVFITDatabase native
    Set source = native.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
    stage = "Check native report identity"
    If source.EOF Then Err.Raise vbObjectError + 5313, , "Returned report is missing; deletion is never applied automatically."
    If source!ReportID <> returnContext.External Then Err.Raise vbObjectError + 5313, , "Returned native identity differs. Import as a new report or supply the original exported file."
    Set returnContext.Values = NAVFITReadFields(source, returnContext.Century)
    Set returnContext.NativeBase = CreateObject("Scripting.Dictionary"): Set returnContext.NativeIncoming = CreateObject("Scripting.Dictionary")
    Set mappedNames = CreateObject("Scripting.Dictionary"): mappedNames.CompareMode = vbTextCompare
    For Each row In NAVFITMapRows(returnContext.Values("FormKind"))
        parts = Split(row, "|"): mappedNames.Add CStr(parts(1)), True
    Next
    Set frozenNative = baseline.OpenRecordset("SELECT * FROM Reports WHERE ReportID=" & returnContext.External, dbOpenSnapshot)
    If frozenNative.EOF Then Err.Raise vbObjectError + 5313, , "Frozen native report is missing."
    For Each nativeField In source.Fields
        If Not mappedNames.Exists(nativeField.Name) Then
            returnContext.NativeBase.Add nativeField.Name, frozenNative.Fields(nativeField.Name).Value
            returnContext.NativeIncoming.Add nativeField.Name, nativeField.Value
        End If
    Next
    frozenNative.Close: Set frozenNative = Nothing
    If returnContext.Values("FormKind") <> baseValues("FormKind") Then Err.Raise vbObjectError + 5313, , "Returned report type changed."
    source.MoveNext
    If Not source.EOF Then Err.Raise vbObjectError + 5313, , "Single-report return contains additional reports. Import those separately; identity is not inferred from SSN."
    source.Close: Set source = Nothing: native.Close: Set native = Nothing: baseline.Close: Set baseline = Nothing
    Set returnContext.Frozen = baseValues
    Set returnContext.Decisions = CreateObject("Scripting.Dictionary")
    For Each row In NAVFITMapRows(baseValues("FormKind"))
        parts = Split(row, "|"): key = parts(0)
        outcome = NAVFITMergeDecision(baseValues(key), returnContext.Before(key), returnContext.Values(key))
        returnContext.Decisions.Add key, outcome
        Set fields = Nothing
        result.Add key & "|" & outcome
    Next
    For Each row In returnContext.NativeBase.Keys
        outcome = "Unchanged native value"
        If Not SameFieldValue(returnContext.NativeBase(row), returnContext.NativeIncoming(row)) Then outcome = "Accept native change"
        result.Add "Native:" & row & "|" & outcome
    Next
    If SourceFileHash(path) <> returnContext.Hash Or SourceFileHash(baselinePath) <> returnContext.BaselineHash Then Err.Raise vbObjectError + 5314, , "File changed during preview."
    returnContext.Path = path: returnContext.Baseline = baselinePath: returnContext.StageGuid = resumeToken
    If Not restoredChoices Is Nothing Then
        If restoredChoices.Count <> returnContext.Decisions.Count + 1 Then Err.Raise vbObjectError + 9202, , "Saved return choices are incomplete. Cancel or preview again."
        For Each row In returnContext.Decisions.Keys
            If Not restoredChoices.Exists(CStr(row)) Then Err.Raise vbObjectError + 9202, , "Saved return choices no longer match this mapping."
            outcome = CStr(restoredChoices(row))
            If outcome <> returnContext.Decisions(row) And outcome <> "Keep current" And outcome <> "Accept returned" Then Err.Raise vbObjectError + 9202, , "Saved return choice is invalid."
            returnContext.Decisions(row) = outcome
        Next
    End If
    NAVFITCheckpointReturned
    Set NAVFITPreviewReturned = result: Exit Function
Failed:
    failure = Err.Number: message = stage & ": " & Err.Description
    On Error Resume Next
    NAVFITClearReturned False
    If Not frozenNative Is Nothing Then frozenNative.Close
    If Not source Is Nothing Then source.Close
    If Not fields Is Nothing Then fields.Close
    If Not meta Is Nothing Then meta.Close
    If Not native Is Nothing Then native.Close
    If Not baseline Is Nothing Then baseline.Close
    On Error GoTo 0
    Err.Raise failure, "NAVFITPreviewReturned", message
End Function
Public Sub NAVFITResolveReturned(ByVal key As String, ByVal choice As String)
    Dim before As String, failure As Long, message As String
    RequireReturnedContext
    If returnContext.Decisions Is Nothing Then Err.Raise vbObjectError + 5315, , "Preview a returned report first."
    If Not returnContext.Decisions.Exists(key) Then Err.Raise vbObjectError + 5315, , "Unknown returned field."
    If choice <> "Keep current" And choice <> "Accept returned" Then Err.Raise vbObjectError + 5315, , "Choose Keep current or Accept returned."
    before = returnContext.Decisions(key)
    On Error GoTo Failed
    returnContext.Decisions(key) = choice
    NAVFITCheckpointReturned
    Exit Sub
Failed:
    failure = Err.Number: message = Err.Description: returnContext.Decisions(key) = before
    Err.Raise failure, "NAVFITResolveReturned", message
End Sub
Private Sub RequireReturnedContext()
    If returnContext.Decisions Is Nothing Then Err.Raise vbObjectError + 5315, , "Preview ownership was reset. Resume the saved preview before continuing."
    WizardPreviewCheck "NAVFITReturn", returnContext.StageGuid
End Sub
Private Sub NAVFITCheckpointReturned()
    Dim choices As Object, key As Variant
    Set choices = CreateObject("Scripting.Dictionary")
    For Each key In returnContext.Decisions.Keys: choices.Add CStr(key), returnContext.Decisions(key): Next
    choices.Add "__Century", CStr(returnContext.Century)
    returnContext.StageGuid = WizardPreviewSave("NAVFITReturn", returnContext.Path, returnContext.Hash, returnContext.Baseline, returnContext.BaselineHash, returnContext.Report, returnContext.Revision, choices, returnContext.StageGuid)
End Sub
Public Function NAVFITResumeReturned() As Collection
    Dim saved As Object, rows As Collection, key As Variant
    Set saved = WizardPreviewLoad("NAVFITReturn")
    WizardPreviewCheck "NAVFITReturn", saved("StageGuid")
    Set rows = NAVFITPreviewReturned(saved("SourcePath"), saved("BaselinePath"), CLng(saved("Choices")("__Century")), saved("StageGuid"), saved("Choices"))
    Set NAVFITResumeReturned = NAVFITReturnedRows()
End Function
Public Function NAVFITReturnedRows() As Collection
    Dim rows As New Collection, key As Variant
    RequireReturnedContext
    For Each key In returnContext.Decisions.Keys: rows.Add CStr(key) & "|" & returnContext.Decisions(key): Next
    For Each key In returnContext.NativeBase.Keys: rows.Add "Native:" & CStr(key) & "|" & NAVFITReturnedDecision("Native:" & CStr(key)): Next
    Set NAVFITReturnedRows = rows
End Function
Private Function NAVFITHasTable(ByVal db As DAO.Database, ByVal name As String) As Boolean
    Dim table As DAO.TableDef
    For Each table In db.TableDefs
        If table.Name = name Then NAVFITHasTable = True: Exit Function
    Next
End Function
Public Function NAVFITReturnedDecision(ByVal key As String) As String
    RequireReturnedContext
    If returnContext.Decisions Is Nothing Then Err.Raise vbObjectError + 5315, , "Preview first."
    If Left$(key, 7) = "Native:" Then
        key = Mid$(key, 8)
        NAVFITReturnedDecision = "Unchanged native value"
        If Not SameFieldValue(returnContext.NativeBase(key), returnContext.NativeIncoming(key)) Then NAVFITReturnedDecision = "Accept native change"
    Else
        NAVFITReturnedDecision = returnContext.Decisions(key)
    End If
End Function
Public Function NAVFITApplyReturned() As Long
    Dim changes As Object, key As Variant, id As Long, w As DAO.Workspace, trans As Boolean, failure As Long, message As String
    RequireReturnedContext
    If returnContext.Decisions Is Nothing Then Err.Raise vbObjectError + 5315, , "Preview a returned report first."
    NAVFITCheckTarget returnContext.Report
    If CurrentRevision(returnContext.Report) <> returnContext.Revision Then Err.Raise vbObjectError + 5316, , "Swift report changed after preview; preview again."
    If SourceFileHash(returnContext.Path) <> returnContext.Hash Or SourceFileHash(returnContext.Baseline) <> returnContext.BaselineHash Then Err.Raise vbObjectError + 5314, , "Exchange file changed after preview."
    Set changes = CreateObject("Scripting.Dictionary")
    For Each key In returnContext.Decisions.Keys
        If returnContext.Decisions(key) = "Conflict" Then Err.Raise vbObjectError + 5317, , "Resolve every conflicting field before Apply."
        If returnContext.Decisions(key) = "Accept returned" Then changes.Add key, returnContext.Values(key)
    Next
    On Error GoTo Failed
    Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    Call OpenReviewFile(returnContext.Path, False)
    id = ImportReviewedNarrative(returnContext.External, "", returnContext.Report, returnContext.Revision, returnContext.Hash, True, returnContext.Century, changes)
    Call CloseReviewFile
    Call WizardPreviewCancel("NAVFITReturn", returnContext.StageGuid)
    w.CommitTrans: trans = False
    NAVFITClearReturned False: NAVFITApplyReturned = id: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    Call CloseReviewFile
    If trans Then w.Rollback
    On Error GoTo 0
    Err.Raise failure, "NAVFITApplyReturned", message
End Function
Public Function NAVFITClearReturned(Optional ByVal cancelSaved As Boolean = True) As Boolean
    If cancelSaved Then Call WizardPreviewCancel("NAVFITReturn")
    returnContext.Path = "": returnContext.Hash = "": returnContext.Baseline = "": returnContext.BaselineHash = "": returnContext.Report = 0: returnContext.Revision = 0: returnContext.External = 0: returnContext.Century = 0
    Set returnContext.Values = Nothing: Set returnContext.Decisions = Nothing: Set returnContext.Before = Nothing: Set returnContext.Frozen = Nothing: returnContext.NativeText = ""
    Set returnContext.NativeBase = Nothing: Set returnContext.NativeIncoming = Nothing
    returnContext.StageGuid = ""
    NAVFITClearReturned = True
End Function

Public Function NAVFITReturnedDetails(ByVal key As String) As String
    RequireReturnedContext
    If returnContext.Decisions Is Nothing Then Err.Raise vbObjectError + 5315, , "Preview first."
    If Left$(key, 7) = "Native:" Then
        key = Mid$(key, 8)
        NAVFITReturnedDetails = "Native-only column: " & key & vbCrLf & "Exported/current native value: " & NAVFITDisplay(returnContext.NativeBase(key)) & vbCrLf & "Returned native value: " & NAVFITDisplay(returnContext.NativeIncoming(key)) & vbCrLf & "Retained as native source content when Apply is chosen. It is not a Swift editor field or local approval. Cancel this return if these source changes are unwanted."
        Exit Function
    End If
    NAVFITReturnedDetails = UXFieldCaption(key, returnContext.Before("FormKind")) & vbCrLf & "Exported baseline: " & NAVFITDisplay(returnContext.Frozen(key)) & vbCrLf & "Current Swift: " & NAVFITDisplay(returnContext.Before(key)) & vbCrLf & "Returned NAVFIT: " & NAVFITDisplay(returnContext.Values(key)) & vbCrLf & "Decision: " & returnContext.Decisions(key)
End Function
Public Function NAVFITDisplay(ByVal value As Variant) As String
    If IsNull(value) Then
        NAVFITDisplay = "[null]"
    ElseIf VarType(value) = vbDate Then
        NAVFITDisplay = Format$(value, "yyyy-mm-dd")
    ElseIf CStr(value) = "" Then
        NAVFITDisplay = "[empty text]"
    Else
        NAVFITDisplay = CStr(value)
    End If
End Function

Public Sub NAVFITRebuildSource(ByVal receiptID As Long, ByVal blankTemplate As String, ByVal destination As String)
    Dim db As DAO.Database, localDb As DAO.Database, rows As DAO.Recordset, cells As DAO.Recordset, table As Variant, q As DAO.QueryDef, names As String, parameterNames As String, i As Long, value As Variant, decimalMark As String
    Dim failure As Long, message As String, stage As String, inserted As DAO.Recordset, rowID As Long, identityField As String, seenFields As Object
    On Error GoTo Failed
    decimalMark = Mid$(CStr(1.5), 2, 1)
    If InStr(Nz(MajorScalar("SELECT ValueEncoding FROM ImportReceipt WHERE ImportReceiptID=" & receiptID), ""), "decimal=" & decimalMark) = 0 Then Err.Raise vbObjectError + 5320, , "Retained values use another decimal encoding. Restore the original source or convert its provenance explicitly."
    stage = "Clone template": FileCopy blankTemplate, destination
    Set db = DBEngine.OpenDatabase(destination): ValidateNAVFITDatabase db: Set localDb = CurrentDb
    For Each table In Array("Reports", "Folders", "Summary")
        stage = "Clear table " & table: db.Execute "DELETE FROM [" & table & "]", dbFailOnError
        Set rows = localDb.OpenRecordset("SELECT DISTINCT SourceRow FROM ImportRawValue WHERE ImportReceiptID=" & receiptID & " AND SourceTable=" & QuoteSQL(CStr(table)) & " ORDER BY SourceRow", dbOpenSnapshot)
        Do While Not rows.EOF
            Set cells = localDb.OpenRecordset("SELECT * FROM ImportRawValue WHERE ImportReceiptID=" & receiptID & " AND SourceTable=" & QuoteSQL(CStr(table)) & " AND SourceRow=" & rows!SourceRow & " ORDER BY FieldOrdinal", dbOpenSnapshot)
            names = "": parameterNames = "": i = 0: Set seenFields = CreateObject("Scripting.Dictionary")
            Do While Not cells.EOF
                If i > 0 Then names = names & ",": parameterNames = parameterNames & ","
                ' The official v33 upgrader appends columns. Retain source ordinals for
                ' provenance, but bind verified export columns by unique name/type.
                If cells!FieldOrdinal <> i Or seenFields.Exists(CStr(cells!FieldKey)) Then Err.Raise vbObjectError + 5321, , "Duplicate or incomplete native provenance."
                If Not NAVFITHasField(db.TableDefs(CStr(table)), CStr(cells!FieldKey)) Then Err.Raise vbObjectError + 5321, , "Unknown retained native field."
                If cells!DAOType <> db.TableDefs(CStr(table)).Fields(CStr(cells!FieldKey)).Type Then Err.Raise vbObjectError + 5321, , "Retained native type does not match the verified template."
                seenFields.Add CStr(cells!FieldKey), True
                names = names & "[" & cells!FieldKey & "]"
                ' DAO infers an empty text parameter as Null. A constant empty
                ' string preserves the recorded distinction without interpolation.
                If cells!DAOType = dbMemo Then
                    parameterNames = parameterNames & "Null"
                ElseIf Not cells!IsNullValue And cells!DAOType = dbText And Len(Nz(cells!RawValue, "")) = 0 Then
                    parameterNames = parameterNames & "''"
                Else
                    parameterNames = parameterNames & "p" & i
                End If
                i = i + 1: cells.MoveNext
            Loop
            If i <> db.TableDefs(CStr(table)).Fields.Count Then Err.Raise vbObjectError + 5321, , "Incomplete native provenance row."
            Set q = db.CreateQueryDef("", "INSERT INTO [" & table & "] (" & names & ") VALUES (" & parameterNames & ")")
            cells.MoveFirst: i = 0
            Do While Not cells.EOF
                stage = "Bind " & table & " " & cells!FieldKey & " type " & cells!DAOType
                value = Null
                If Not cells!IsNullValue Then
                    Select Case CLng(cells!DAOType)
                        Case dbBoolean: value = CBool(cells!RawValue)
                        Case dbInteger: value = CInt(cells!RawValue)
                        Case dbLong: value = CLng(cells!RawValue)
                        Case dbDate: value = CDate(CDbl(cells!RawValue))
                        Case dbDecimal: value = CDec(cells!RawValue)
                        Case dbText, dbMemo: value = Nz(cells!RawValue, "")
                        Case Else: Err.Raise vbObjectError + 5321, , "Unsupported retained native type."
                    End Select
                End If
                If cells!FieldKey = "ReportID" Or cells!FieldKey = "FolderID" Then rowID = CLng(value): identityField = cells!FieldKey
                If cells!DAOType <> dbMemo And Not (Not cells!IsNullValue And cells!DAOType = dbText And Len(Nz(cells!RawValue, "")) = 0) Then q.parameters("p" & i).Value = value
                i = i + 1: cells.MoveNext
            Loop
            stage = "Insert " & table: q.Execute dbFailOnError: q.Close: Set q = Nothing
            ' Memo parameters are inferred as short text by ACE. Restore full
            ' typed Memo content through the inserted row, including Null/empty.
            If CStr(table) <> "Summary" Then
                Set inserted = db.OpenRecordset("SELECT * FROM [" & table & "] WHERE [" & identityField & "]=" & rowID, dbOpenDynaset)
                inserted.Edit: cells.MoveFirst
                Do While Not cells.EOF
                    If cells!DAOType = dbMemo Then
                        value = Null
                        If Not cells!IsNullValue Then value = Nz(cells!RawValue, "")
                        inserted.Fields(CStr(cells!FieldKey)).Value = value
                    End If
                    cells.MoveNext
                Loop
                inserted.Update: inserted.Close: Set inserted = Nothing
            End If
            cells.Close: Set cells = Nothing: rows.MoveNext
        Loop
        rows.Close: Set rows = Nothing
    Next
    db.Close: Exit Sub
Failed:
    failure = Err.Number: message = stage & ": " & Err.Description
    On Error Resume Next
    If Not q Is Nothing Then q.Close
    If Not inserted Is Nothing Then inserted.Close
    If Not cells Is Nothing Then cells.Close
    If Not rows Is Nothing Then rows.Close
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    Err.Raise failure, "NAVFITRebuildSource", message
End Sub





Public Function NAVFITReturnedCaption(ByVal key As String) As String
    If Left$(key, 7) = "Native:" Then
        NAVFITReturnedCaption = "Native-only: " & Mid$(key, 8)
    Else
        NAVFITReturnedCaption = UXFieldCaption(key, returnContext.Before("FormKind"))
    End If
End Function
