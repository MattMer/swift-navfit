Attribute VB_Name = "modImportReview"
Option Compare Database
Option Explicit

Private reviewDb As DAO.Database
Private reviewReader As DAO.Workspace
Private reviewPrefix As String
Private reviewOriginalPath As String, reviewOriginalHash As String
Private reviewCopyPath As String, reviewPreparedHash As String

Public Function OpenReviewFile(ByVal filePath As String, Optional ByVal showReview As Boolean = True) As Boolean
    Dim candidate As DAO.Database, reader As DAO.Workspace, r As DAO.Recordset
    Dim i As Long, prefix As String, kind As String, failure As Long, message As String
    Dim prepared As String, originalHash As String, preparedHash As String
    On Error GoTo Failed
    If StrComp(filePath, CurrentProject.FullName, vbTextCompare) = 0 Then Err.Raise vbObjectError + 400, , "Choose a NAVFIT source or isolated staging file."
    Set reader = DBEngine.CreateWorkspace("SwiftReview", "admin", "", dbUseJet)
    Set candidate = reader.OpenDatabase(filePath, False, True)
    prefix = ""
    On Error Resume Next
    prefix = candidate.TableDefs("RawReports").Name
    On Error GoTo Failed
    If prefix = "RawReports" Then prefix = "Raw" Else prefix = ""
    Call ValidateReviewSchema(candidate, prefix, prefix = "")
    If prefix = "" Then
        If NAVFITNeedsV33Upgrade(candidate) Then
            originalHash = SourceFileHash(filePath)
            candidate.Close: Set candidate = Nothing
            prepared = Environ$("TEMP") & "\SwiftNAVFIT-v33p2-" & NewIdentity() & ".accdb"
            FileCopy filePath, prepared: SetAttr prepared, vbNormal
            If SourceFileHash(prepared) <> originalHash Or SourceFileHash(filePath) <> originalHash Then Err.Raise vbObjectError + 420, , "Source changed while preparing the upgrade. Choose it again."
            UpgradeNAVFITV33Copy prepared, reader
            preparedHash = SourceFileHash(prepared)
            Set candidate = reader.OpenDatabase(prepared, False, True)
            Call ValidateReviewSchema(candidate, prefix)
        End If
    End If
    Set r = candidate.OpenRecordset("SELECT ReportID,ReportType,FullName FROM [" & prefix & "Reports] ORDER BY ReportID", dbOpenSnapshot)
    Call CloseReviewFile()
    Set reviewDb = candidate: Set reviewReader = reader: reviewPrefix = prefix
    reviewOriginalPath = filePath: reviewOriginalHash = originalHash
    reviewCopyPath = prepared: reviewPreparedHash = preparedHash
    If Not showReview Then r.Close: OpenReviewFile = True: Exit Function
    Call OpenScreen("frmImportReview")
    With Forms!frmImportReview
        !txtSource = filePath
        !txtSource.ControlTipText = ReviewUpgradeMessage()
        !lstImportReports.RowSourceType = "Value List"
        !lstImportReports.RowSource = ""
        Do While Not r.EOF
            'Only structural identifiers and fixed classifications enter the list.
            kind = "Unknown"
            If StrComp(Nz(r!ReportType, ""), "Eval", vbBinaryCompare) = 0 Then kind = "EVAL"
            If StrComp(Nz(r!ReportType, ""), "Chief", vbBinaryCompare) = 0 Then kind = "CHIEFEVAL"
            If StrComp(Nz(r!ReportType, ""), "FitRep", vbBinaryCompare) = 0 Then kind = "FITREP"
            !lstImportReports.AddItem CStr(r!ReportID) & ";" & kind & ";" & IIf(kind = "Unknown", "Quarantined", "Pending mapping review") & ";" & Replace(Replace(Replace(Nz(r!FullName, "Unnamed source member"), ";", ","), vbCr, " "), vbLf, " ")
            r.MoveNext
        Loop
        !txtFieldValue = "Select a report, then inspect its Comments field before creating a narrative draft."
        !lblImportResult.Caption = ReviewUpgradeMessage() & "No import performed."
    End With
    Call UXNarrativeInputChanged
    r.Close: OpenReviewFile = True: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not r Is Nothing Then r.Close
    If Not candidate Is Nothing Then candidate.Close
    If Not reader Is Nothing Then reader.Close
    If Len(prepared) > 0 Then Kill prepared
    On Error GoTo 0
    Err.Raise failure, "OpenReviewFile", message
End Function


' Keep receipts/duplicates bound to the original bytes, not a random temporary name.
' Recheck both original and prepared input before preview/apply to retain stale guards.
Private Function ReviewImportHash() As String
    Dim hash As String
    hash = SourceFileHash(reviewOriginalPath)
    If Len(reviewCopyPath) > 0 Then
        If hash <> reviewOriginalHash Or SourceFileHash(reviewCopyPath) <> reviewPreparedHash Then Err.Raise vbObjectError + 420, , "Source or prepared upgrade changed. Choose the source again."
    End If
    ReviewImportHash = hash
End Function
' Diagnostic path of the owned disposable conversion; empty for current/staged input.
Public Function ReviewPreparedCopyPath() As String
    ReviewPreparedCopyPath = reviewCopyPath
End Function
Public Function ReviewUpgradeMessage() As String
    If Len(reviewCopyPath) > 0 Then ReviewUpgradeMessage = "Older v33 database upgraded automatically for import to Patch 2 on a temporary copy. Original unchanged. "
End Function

Public Function ChooseReviewFile() As Boolean
    On Error GoTo Failed
    With Application.FileDialog(3)
        .Title = "Select a NAVFIT or staged database for read-only review"
        .AllowMultiSelect = False
        .Filters.Clear: .Filters.Add "Access source database", "*.accdb"
        If .Show <> -1 Then Exit Function
        ChooseReviewFile = OpenReviewFile(.SelectedItems(1))
    End With
    Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Open staging review"
End Function

Public Function ReviewReportFields() As Boolean
    Dim r As DAO.Recordset, i As Long, id As Long
    If reviewDb Is Nothing Then Exit Function
    id = Nz(Forms!frmImportReview!lstImportReports, 0)
    If id = 0 Then Exit Function
    Set r = reviewDb.OpenRecordset("SELECT * FROM [" & reviewPrefix & "Reports] WHERE ReportID=" & id, dbOpenSnapshot)
    With Forms!frmImportReview!lstImportFields
        .RowSourceType = "Value List": .RowSource = ""
        If Not r.EOF Then
            For i = 0 To r.Fields.Count - 1
                .AddItem CStr(i) & ";" & r.Fields(i).Name & ";" & CStr(r.Fields(i).Type) & ";" & IIf(IsNull(r.Fields(i).Value), "Null", "Value")
            Next
        End If
    End With
    Call UXNarrativeInputChanged
    Forms!frmImportReview!txtFieldValue = "Select a field to inspect its stored value."
    r.Close: ReviewReportFields = True
End Function

Public Function ReviewFieldValue() As Boolean
    Dim r As DAO.Recordset, id As Long, index As Long, value As Variant
    If reviewDb Is Nothing Then Exit Function
    If IsNull(Forms!frmImportReview!lstImportFields) Then Exit Function
    id = Nz(Forms!frmImportReview!lstImportReports, 0)
    index = CLng(Forms!frmImportReview!lstImportFields)
    Set r = reviewDb.OpenRecordset("SELECT * FROM [" & reviewPrefix & "Reports] WHERE ReportID=" & id, dbOpenSnapshot)
    If Not r.EOF And index >= 0 And index < r.Fields.Count Then
        value = r.Fields(index).Value
        If IsNull(value) Then
            Forms!frmImportReview!txtFieldValue = "[NULL]"
        ElseIf VarType(value) = vbDate Then
            Forms!frmImportReview!txtFieldValue = Format$(value, "yyyy-mm-dd hh:nn:ss")
        Else
            Forms!frmImportReview!txtFieldValue = CStr(value)
        End If
    End If
    r.Close: ReviewFieldValue = True
End Function

Public Function CloseReviewFile() As Boolean
    If Not reviewDb Is Nothing Then reviewDb.Close
    If Not reviewReader Is Nothing Then reviewReader.Close
    Set reviewDb = Nothing: Set reviewReader = Nothing
    If Len(reviewCopyPath) > 0 Then
        If Len(Dir$(reviewCopyPath)) > 0 Then Kill reviewCopyPath
    End If
    reviewCopyPath = "": reviewPreparedHash = "": reviewOriginalPath = "": reviewOriginalHash = ""
    CloseReviewFile = True
End Function

Public Function LeaveImportReview() As Boolean
    Call CloseReviewFile()
    DoCmd.Close acForm, "frmImportReview", acSaveNo
    Call OpenScreen("frmHome")
    LeaveImportReview = True
End Function

Public Function ImportReviewedNarrative(ByVal externalID As Long, Optional ByVal memberDisplay As String = "", Optional ByVal destinationID As Long = 0, Optional ByVal expectedRevision As Long = 0, Optional ByVal expectedSourceHash As String = "", Optional ByVal fullReport As Boolean = False, Optional ByVal counselingCentury As Long = 0, Optional ByVal acceptedChanges As Variant) As Long
    Dim r As DAO.Recordset, receipt As DAO.Recordset, raw As DAO.Recordset, values As DAO.Recordset
    Dim db As DAO.Database, w As DAO.Workspace, hash As String, kind As String, id As Long, receiptID As Long, operationID As Long, beforeRevision As Variant, target As DAO.Recordset
    Dim mapped As Object
    Dim table As Variant, row As Long, i As Long, value As Variant, trans As Boolean, failure As Long, message As String
    If reviewDb Is Nothing Then Err.Raise vbObjectError + 410, , "Open and review a source database first."
    If Len(memberDisplay) > 120 Then Err.Raise vbObjectError + 411, , "Member display name exceeds 120 characters."
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0)
    On Error GoTo Failed
    operationID = BeginOperation("Narrative import", ReviewUpgradeMessage() & reviewOriginalPath & " report " & externalID)
    hash = ReviewImportHash()
    If Len(expectedSourceHash) > 0 And hash <> expectedSourceHash Then Err.Raise vbObjectError + 420, , "Source changed after preview. Choose the source and preview again."
    If DCount("*", "ImportReceipt", "SourceHash='" & hash & "' AND ExternalReportID=" & externalID) > 0 Then Err.Raise vbObjectError + 412, , "This exact source report has already been imported. Existing drafts are not overwritten."
    Set r = reviewDb.OpenRecordset("SELECT * FROM [" & reviewPrefix & "Reports] WHERE ReportID=" & externalID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 413, , "Source report does not exist."
    kind = ""
    If StrComp(Nz(r!ReportType, ""), "Eval", vbBinaryCompare) = 0 Then kind = "EVAL"
    If StrComp(Nz(r!ReportType, ""), "Chief", vbBinaryCompare) = 0 Then kind = "CHIEFEVAL"
    If StrComp(Nz(r!ReportType, ""), "FitRep", vbBinaryCompare) = 0 Then kind = "FITREP"
    If kind = "" Then Err.Raise vbObjectError + 414, , "Unknown or null report type is quarantined."
    If fullReport Then Set mapped = NAVFITReadFields(r, counselingCentury)
    w.BeginTrans: trans = True
    beforeRevision = Null
    If destinationID > 0 Then
        If CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 416, , "Close the editor before applying an incoming narrative."
        Set target = db.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & destinationID, dbOpenSnapshot)
        If target.EOF Then Err.Raise vbObjectError + 417, , "Destination draft missing."
        If target!FormKind <> kind Then Err.Raise vbObjectError + 418, , "Source and destination report types differ."
        If target!CurrentRevisionID <> expectedRevision Then Err.Raise vbObjectError + 419, , "Destination changed after preview. Compare again."
        beforeRevision = target!CurrentRevisionID
        If fullReport Then
            If Not IsObject(acceptedChanges) Then Err.Raise vbObjectError + 5111, , "Preview and resolve the returned file against its export baseline first."
            NAVFITCheckTarget destinationID
            id = CommitSavedFields(destinationID, expectedRevision, acceptedChanges, "Accept returned NAVFIT report")
        Else
        id = CommitDraft(destinationID, kind, Nz(target!Title, ""), Nz(target!MemberName, ""), target!PeriodStart, target!PeriodEnd, Nz(target!DraftNotes, ""), "Accept incoming narrative", r!Comments)
        End If
        target.Close
    Else
        If fullReport Then
            If memberDisplay = "" Then memberDisplay = Nz(r!FullName, "Unnamed imported member")
            id = NAVFITCommitFields(0, mapped, memberDisplay, "Full NAVFIT report import")
        Else
            id = CommitDraft(0, kind, "Imported narrative " & externalID, memberDisplay, Null, Null, "Other official fields require mapping and review. Source values retained in import provenance.", "Import confirmed narrative", r!Comments)
        End If
    End If
    r.Close
    Set target = db.OpenRecordset("SELECT CurrentRevisionID FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Set receipt = db.OpenRecordset("ImportReceipt", dbOpenDynaset)
    receipt.AddNew: receipt!ReceiptGuid = NewIdentity(): receipt!SourceHash = hash
    receipt!SourcePath = reviewOriginalPath: receipt!ExternalReportID = externalID: receipt!ReportID = id
    receipt!AuthorProfileID = ActiveContributorID()
    receipt!ImportedAt = Now(): receipt!AdapterVersion = "comments-v1/local-CStr"
    If fullReport Then receipt!AdapterVersion = "navfit-full-v1/local-CStr"
    receipt!BeforeRevisionID = beforeRevision: receipt!AppliedRevisionID = target!CurrentRevisionID
    target.Close
    receipt!OperationID = operationID: receipt!IsUndone = False
    receipt!ValueEncoding = "VBA CStr; dates=OLE double; decimal=" & Mid$(CStr(1.5), 2, 1)
    receipt.Update: receipt.Bookmark = receipt.LastModified: receiptID = receipt!ImportReceiptID: receipt.Close
    Set values = db.OpenRecordset("ImportRawValue", dbOpenDynaset)
    For Each table In Array("Folders", "Reports", "Summary")
        Set raw = reviewDb.OpenRecordset("SELECT * FROM [" & reviewPrefix & table & "]", dbOpenSnapshot)
        row = 0
        Do While Not raw.EOF
            For i = 0 To raw.Fields.Count - 1
                value = raw.Fields(i).Value
                values.AddNew: values!ImportReceiptID = receiptID: values!SourceTable = table
                values!SourceRow = row: values!FieldOrdinal = i: values!FieldKey = raw.Fields(i).Name
                values!DAOType = raw.Fields(i).Type: values!IsNullValue = IsNull(value)
                If Not IsNull(value) Then
                    If VarType(value) = vbDate Then values!RawValue = CStr(CDbl(value)) Else values!RawValue = CStr(value)
                End If
                values.Update
            Next
            row = row + 1: raw.MoveNext
        Loop
        raw.Close
    Next
    values.Close
    If ReviewImportHash() <> hash Then Err.Raise vbObjectError + 415, , "Source changed during import. Review it again."
    If IsNull(beforeRevision) Or Nz(beforeRevision, 0) <> CurrentRevision(id) Then AuditCorrelateRevision CurrentRevision(id), operationID
    Call FinishOperation(operationID, "Completed", ReviewUpgradeMessage() & "Applied narrative to Swift report " & id)
    w.CommitTrans: trans = False
    ImportReviewedNarrative = id
    Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If trans Then w.Rollback
    If Not r Is Nothing Then r.Close
    If Not raw Is Nothing Then raw.Close
    If Not values Is Nothing Then values.Close
    If Not receipt Is Nothing Then receipt.Close
    If Not target Is Nothing Then target.Close
    If operationID > 0 Then FinishOperation operationID, "Failed", message
    On Error GoTo 0
    Err.Raise failure, "ImportReviewedNarrative", message
End Function

Public Function ImportSelectedNarrativeUI() As Boolean
    ImportSelectedNarrativeUI = UXApplyNarrativeImport()
End Function
Public Function NarrativeImportValues(ByVal externalID As Long) As Object
    Dim r As DAO.Recordset, result As Object, kind As String, sourceHash As String
    If reviewDb Is Nothing Then Err.Raise vbObjectError + 410, , "Choose a source database first."
    sourceHash = ReviewImportHash()
    Set r = reviewDb.OpenRecordset("SELECT ReportType,FullName,Comments FROM [" & reviewPrefix & "Reports] WHERE ReportID=" & externalID, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 413, , "Select a source report."
    Select Case Nz(r!ReportType, "")
        Case "Eval": kind = "EVAL"
        Case "Chief": kind = "CHIEFEVAL"
        Case "FitRep": kind = "FITREP"
        Case Else: r.Close: Err.Raise vbObjectError + 414, , "This source report is quarantined because its type is unsupported."
    End Select
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "FormKind", kind: result.Add "OfficialComments", r!Comments.Value: result.Add "SourceName", r!FullName.Value
    r.Close
    If ReviewImportHash() <> sourceHash Then Err.Raise vbObjectError + 420, , "Source changed during preview. Choose it again."
    result.Add "SourceHash", sourceHash
    Set NarrativeImportValues = result
End Function

Private Sub ValidateReviewSchema(ByVal db As DAO.Database, ByVal prefix As String, Optional ByVal allowV33 As Boolean = False)
    ValidateNAVFITSourceSchema db, prefix, allowV33
End Sub





Public Function PreviewIncomingNarrative(ByVal externalID As Long, ByVal reportID As Long) As Boolean
    Dim source As DAO.Recordset, target As DAO.Recordset
    If reviewDb Is Nothing Then Err.Raise vbObjectError + 420, , "Open a source first."
    Set source = reviewDb.OpenRecordset("SELECT * FROM [" & reviewPrefix & "Reports] WHERE ReportID=" & externalID, dbOpenSnapshot)
    Set target = CurrentDb.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & reportID, dbOpenSnapshot)
    If source.EOF Or target.EOF Then Err.Raise vbObjectError + 421, , "Select source and destination reports."
    DoCmd.OpenForm "frmReconcile"
    With Forms!frmReconcile
        !txtExternalID = externalID: !txtReportID = reportID: !txtExpectedRevision = target!CurrentRevisionID
        !txtSourceHash = ReviewImportHash()
        !txtCurrent = target!OfficialComments: !txtIncoming = source!Comments
        !lblContext.Caption = "Swift report " & reportID & " - saved revision " & target!CurrentRevisionID & " versus source report " & externalID
    End With
    source.Close: target.Close: Forms!frmReconcile.SetFocus: PreviewIncomingNarrative = True
End Function
Public Function PreviewReconcileUI() As Boolean
    On Error GoTo Failed
    PreviewReconcileUI = PreviewIncomingNarrative(Nz(Forms!frmImportReview!lstImportReports, 0), Nz(Forms!frmImportReview!cboDestinationReport, 0)): Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Compare incoming narrative"
End Function
Public Function AcceptIncomingNarrative(ByVal externalID As Long, ByVal reportID As Long, ByVal expectedRevision As Long, ByVal expectedHash As String) As Long
    If reviewDb Is Nothing Then Err.Raise vbObjectError + 420, , "Source review is closed."
    If ReviewImportHash() <> expectedHash Then Err.Raise vbObjectError + 422, , "Source changed after preview. Compare again."
    AcceptIncomingNarrative = ImportReviewedNarrative(externalID, "", reportID, expectedRevision)
End Function
Public Function AcceptIncomingUI() As Boolean
    On Error GoTo Failed
    If MsgBox("Replace only the saved official narrative with the reviewed incoming value? This records a new revision and import provenance. Other Swift fields stay unchanged.", vbOKCancel + vbQuestion) <> vbOK Then Exit Function
    With Forms!frmReconcile
        Call AcceptIncomingNarrative(CLng(!txtExternalID), CLng(!txtReportID), CLng(!txtExpectedRevision), CStr(!txtSourceHash))
    End With
    DoCmd.Close acForm, "frmReconcile", acSaveNo
    If CurrentProject.AllForms("frmLibrary").IsLoaded Then Call RefreshLibrary()
    AcceptIncomingUI = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Accept incoming narrative"
End Function
Public Function CloseReconcileUI() As Boolean
    DoCmd.Close acForm, "frmReconcile", acSaveNo
    If CurrentProject.AllForms("frmImportReview").IsLoaded Then Forms!frmImportReview.SetFocus
    CloseReconcileUI = True
End Function

Public Function FullNAVFITImportValues(ByVal externalID As Long, Optional ByVal counselingCentury As Long = 0) As Object
    Dim r As DAO.Recordset, result As Object, hash As String
    If reviewDb Is Nothing Then Err.Raise vbObjectError + 410, , "Choose a source database first."
    hash = ReviewImportHash()
    Set r = reviewDb.OpenRecordset("SELECT * FROM [" & reviewPrefix & "Reports] WHERE ReportID=" & externalID, dbOpenSnapshot)
    If r.EOF Then r.Close: Err.Raise vbObjectError + 413, , "Select a source report."
    Set result = NAVFITReadFields(r, counselingCentury): r.Close
    If ReviewImportHash() <> hash Then Err.Raise vbObjectError + 420, , "Source changed during preview."
    result.Add "SourceHash", hash
    Set FullNAVFITImportValues = result
End Function
Public Function ImportReviewedReport(ByVal externalID As Long, ByVal expectedSourceHash As String, Optional ByVal counselingCentury As Long = 0, Optional ByVal memberDisplay As String = "") As Long
    ImportReviewedReport = ImportReviewedNarrative(externalID, memberDisplay, 0, 0, expectedSourceHash, True, counselingCentury)
End Function

Public Sub ValidateNAVFITDatabase(ByVal db As DAO.Database)
    Dim t As DAO.TableDef, name As Variant, count As Long
    ValidateReviewSchema db, ""
    For Each t In db.TableDefs
        If Left$(t.Name, 4) <> "MSys" Then
            If t.Name <> "Folders" And t.Name <> "Reports" And t.Name <> "Summary" Then Err.Raise vbObjectError + 5322, , "Native exchange database contains unexpected tables."
            count = count + 1
        End If
    Next
    If count <> 3 Or db.QueryDefs.Count <> 0 Then Err.Raise vbObjectError + 5322, , "Native exchange expects the three verified data tables and no saved queries."
    For Each name In Array("Forms", "Reports", "Modules", "Scripts", "DataAccessPages")
        If db.Containers(name).Documents.Count <> 0 Then Err.Raise vbObjectError + 5322, , "Choose a NAVFIT data database without executable or interface objects."
    Next
End Sub
Public Function NAVFITImportScope(ByVal reportID As Long) As String
    Dim receipt As Variant, adapter As String
    receipt = MajorScalar("SELECT Max(ImportReceiptID) FROM ImportReceipt WHERE ReportID=" & reportID & " AND IsUndone=False")
    If IsNull(receipt) Then Exit Function
    adapter = Nz(DLookup("AdapterVersion", "ImportReceipt", "ImportReceiptID=" & CLng(receipt)), "")
    If adapter = "navfit-full-v1/local-CStr" Then NAVFITImportScope = "Full report" Else NAVFITImportScope = "Narrative only"
End Function

Public Function NAVFITReviewGroupChoices(Optional ByVal century As Long = 0, Optional ByVal gradeOverrides As Object, Optional ByVal discoverOnly As Boolean = False) As Object
    Dim groups As Object, item As Object, r As DAO.Recordset, folder As DAO.Recordset, values As Object, defaults As Object, key As Variant, row As Variant, parts As Variant
    Dim parent As String, fid As Long, kind As String, grade As String, token As String, hash As String, field As DAO.Field, exists As Boolean
    Dim unresolved As Boolean, missingKey As String, blocked As String
    If reviewDb Is Nothing Then Err.Raise vbObjectError + 6940, , "Choose a NAVFIT source first."
    hash = ReviewImportHash(): Set groups = CreateObject("Scripting.Dictionary")
    Set r = reviewDb.OpenRecordset("SELECT * FROM [" & reviewPrefix & "Reports] WHERE ReportType Is Not Null ORDER BY ReportID", dbOpenSnapshot)
    Do While Not r.EOF
        kind = NAVFITKind(r!ReportType): grade = ValidationWornPaygrade(Nz(r!Rate, ""))
        unresolved = (grade = ""): blocked = ""
        parent = Nz(r!Parent, "")
        If Left$(parent, 2) <> "a " Then Err.Raise vbObjectError + 6940, , "Source report has an unsupported parent reference."
        fid = Val(Mid$(parent, 3))
        If parent <> "a " & CStr(fid) Then Err.Raise vbObjectError + 6940, , "Source report has a malformed folder reference."
        missingKey = fid & "|" & kind & "|?"
        If unresolved Then
            grade = "?"
            If Not gradeOverrides Is Nothing Then
                If gradeOverrides.Exists(missingKey) Then grade = CStr(gradeOverrides(missingKey))
            End If
        End If
        If grade = "?" Then
            blocked = "Source rank/rate is blank or unrecognized. Assign a group paygrade before Preview; original report values will be preserved."
        ElseIf Not NativeImportGradeMatches(kind, grade) Then
            blocked = "Source report type and group paygrade disagree. Choose a compatible grade for unresolved source rates, or correct the source report."
        End If
        token = fid & "|" & kind & "|" & grade
        If Not groups.Exists(token) Then
            Set folder = reviewDb.OpenRecordset("SELECT * FROM [" & reviewPrefix & "Folders] WHERE FolderID=" & fid, dbOpenSnapshot)
            If folder.EOF Then Err.Raise vbObjectError + 6940, , "Source report references a missing folder."
            Set item = CreateObject("Scripting.Dictionary"): Set defaults = CreateObject("Scripting.Dictionary")
            item.Add "Title", Nz(folder!FolderName, "Imported group"): item.Add "Kind", kind: item.Add "Paygrade", grade: item.Add "Reports", "": item.Add "Count", 0
            item.Add "UnresolvedCount", 0: item.Add "Blocked", blocked
            For Each row In NAVFITMapRows(kind)
                parts = Split(row, "|"): exists = False
                For Each key In UXSharedKeys()
                    If parts(0) = CStr(key) Then exists = True: Exit For
                Next
                If exists Then
                    exists = False
                    For Each field In folder.Fields
                        If field.Name = parts(1) Then exists = True: Exit For
                    Next
                    If exists And Not discoverOnly Then defaults.Add parts(0), NAVFITDecode(folder.Fields(parts(1)).Value, parts(2), century)
                End If
            Next
            item.Add "Defaults", defaults: item.Add "SourceHash", hash: groups.Add token, item: folder.Close
        End If
        Set item = groups(token)
        If unresolved Then item("UnresolvedCount") = item("UnresolvedCount") + 1
        If item("Reports") <> "" Then item("Reports") = item("Reports") & ","
        item("Reports") = item("Reports") & r!ReportID: item("Count") = item("Count") + 1
        r.MoveNext
    Loop
    r.Close
    If ReviewImportHash() <> hash Then Err.Raise vbObjectError + 6940, , "Source changed during group preview. Choose it again."
    Set NAVFITReviewGroupChoices = groups
End Function
' Only explicit supported paygrades may fill missing source grouping metadata.
Public Function NativeImportGradeMatches(ByVal kind As String, ByVal grade As String) As Boolean
    Dim prefix As String, number As Long
    If Len(grade) < 2 Or Len(grade) > 3 Then Exit Function
    prefix = Left$(grade, 1)
    If Mid$(grade, 2) Like "*[!0-9]*" Then Exit Function
    number = Val(Mid$(grade, 2)): If grade <> prefix & CStr(number) Then Exit Function
    Select Case kind
        Case "EVAL": NativeImportGradeMatches = (prefix = "E" And number >= 1 And number <= 6)
        Case "CHIEFEVAL": NativeImportGradeMatches = (prefix = "E" And number >= 7 And number <= 9)
        Case "FITREP": NativeImportGradeMatches = ((prefix = "O" And number >= 1 And number <= 10) Or (prefix = "W" And number >= 1 And number <= 5))
    End Select
End Function
Public Function ImportNAVFITGroups(ByVal selections As Collection, ByVal expectedHash As String, ByVal century As Long, Optional ByVal gradeOverrides As Object) As Collection
    Dim groups As Object, item As Object, result As New Collection, selected As Variant, ids As Variant, nativeID As Variant, key As Variant, values As Object, defaults As Object
    Dim w As DAO.Workspace, id As Long, g As Long, failure As Long, message As String, trans As Boolean, title As String
    On Error GoTo Failed
    Set groups = NAVFITReviewGroupChoices(century, gradeOverrides)
    If selections.Count = 0 Then Err.Raise vbObjectError + 6941, , "Select at least one source group."
    For Each selected In selections
        If Not groups.Exists(CStr(selected)) Then Err.Raise vbObjectError + 6941, , "Source groups changed. Preview again."
        Set item = groups(CStr(selected))
        If item("Blocked") <> "" Then Err.Raise vbObjectError + 6941, , item("Title") & ": " & item("Blocked")
        If item("SourceHash") <> expectedHash Then Err.Raise vbObjectError + 6941, , "Source changed after preview."
        For Each nativeID In Split(item("Reports"), ",")
            Set values = FullNAVFITImportValues(CLng(nativeID), century)
            If DCount("*", "ImportReceipt", "SourceHash=" & QuoteSQL(expectedHash) & " AND ExternalReportID=" & CLng(nativeID)) > 0 Then Err.Raise vbObjectError + 6941, , "A selected report was already imported. Use returned-file reconciliation for existing drafts."
        Next
    Next
    Set w = DBEngine.Workspaces(0): w.BeginTrans: trans = True
    For Each selected In selections
        Set item = groups(CStr(selected)): title = item("Title") & " - " & item("Kind") & " " & item("Paygrade")
        g = CreateSummaryGroup(title, item("Kind"), item("Paygrade")): Set defaults = item("Defaults")
        For Each key In defaults.Keys: SetFieldDefault "Group", g, CStr(key), defaults(key): Next
        For Each nativeID In Split(item("Reports"), ",")
            id = ImportReviewedReport(CLng(nativeID), expectedHash, century)
            AddGroupMember g, id, item("Paygrade")
        Next
        result.Add g
    Next
    If ReviewImportHash() <> expectedHash Then Err.Raise vbObjectError + 6941, , "Source changed during group import. No selected group was committed."
    w.CommitTrans: trans = False: Set ImportNAVFITGroups = result: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If trans Then w.Rollback
    On Error GoTo 0
    Err.Raise failure, "ImportNAVFITGroups", message
End Function
