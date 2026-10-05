Attribute VB_Name = "modImportReview"
Option Compare Database
Option Explicit

Private reviewDb As DAO.Database
Private reviewReader As DAO.Workspace
Private reviewPrefix As String

Public Function OpenReviewFile(ByVal filePath As String) As Boolean
    Dim candidate As DAO.Database, reader As DAO.Workspace, r As DAO.Recordset
    Dim i As Long, prefix As String, kind As String, failure As Long, message As String
    On Error GoTo Failed
    If StrComp(filePath, CurrentProject.FullName, vbTextCompare) = 0 Then Err.Raise vbObjectError + 400, , "Choose a NAVFIT source or isolated staging file."
    Set reader = DBEngine.CreateWorkspace("SwiftReview", "admin", "", dbUseJet)
    Set candidate = reader.OpenDatabase(filePath, False, True)
    prefix = ""
    On Error Resume Next
    prefix = candidate.TableDefs("RawReports").Name
    On Error GoTo Failed
    If prefix = "RawReports" Then prefix = "Raw" Else prefix = ""
    Call ValidateReviewSchema(candidate, prefix)
    Set r = candidate.OpenRecordset("SELECT ReportID,ReportType,FullName FROM [" & prefix & "Reports] ORDER BY ReportID", dbOpenSnapshot)
    Call CloseReviewFile()
    Set reviewDb = candidate: Set reviewReader = reader: reviewPrefix = prefix
    Call OpenScreen("frmImportReview")
    With Forms!frmImportReview
        !txtSource = filePath
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
        !lblImportResult.Caption = "No import performed."
    End With
    Call UXNarrativeInputChanged
    r.Close: OpenReviewFile = True: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    On Error Resume Next
    If Not candidate Is Nothing Then candidate.Close
    If Not reader Is Nothing Then reader.Close
    On Error GoTo 0
    Err.Raise failure, "OpenReviewFile", message
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
    CloseReviewFile = True
End Function

Public Function LeaveImportReview() As Boolean
    Call CloseReviewFile()
    DoCmd.Close acForm, "frmImportReview", acSaveNo
    Call OpenScreen("frmHome")
    LeaveImportReview = True
End Function

Public Function ImportReviewedNarrative(ByVal externalID As Long, Optional ByVal memberDisplay As String = "", Optional ByVal destinationID As Long = 0, Optional ByVal expectedRevision As Long = 0, Optional ByVal expectedSourceHash As String = "") As Long
    Dim r As DAO.Recordset, receipt As DAO.Recordset, raw As DAO.Recordset, values As DAO.Recordset
    Dim db As DAO.Database, w As DAO.Workspace, hash As String, kind As String, id As Long, receiptID As Long, operationID As Long, beforeRevision As Variant, target As DAO.Recordset
    Dim table As Variant, row As Long, i As Long, value As Variant, trans As Boolean, failure As Long, message As String
    If reviewDb Is Nothing Then Err.Raise vbObjectError + 410, , "Open and review a source database first."
    If Len(memberDisplay) > 120 Then Err.Raise vbObjectError + 411, , "Member display name exceeds 120 characters."
    Set db = CurrentDb: Set w = DBEngine.Workspaces(0)
    On Error GoTo Failed
    operationID = BeginOperation("Narrative import", reviewDb.Name & " report " & externalID)
    hash = SourceFileHash(reviewDb.Name)
    If Len(expectedSourceHash) > 0 And hash <> expectedSourceHash Then Err.Raise vbObjectError + 420, , "Source changed after preview. Choose the source and preview again."
    If DCount("*", "ImportReceipt", "SourceHash='" & hash & "' AND ExternalReportID=" & externalID) > 0 Then Err.Raise vbObjectError + 412, , "This exact source report has already been imported. Existing drafts are not overwritten."
    Set r = reviewDb.OpenRecordset("SELECT * FROM [" & reviewPrefix & "Reports] WHERE ReportID=" & externalID, dbOpenSnapshot)
    If r.EOF Then Err.Raise vbObjectError + 413, , "Source report does not exist."
    kind = ""
    If StrComp(Nz(r!ReportType, ""), "Eval", vbBinaryCompare) = 0 Then kind = "EVAL"
    If StrComp(Nz(r!ReportType, ""), "Chief", vbBinaryCompare) = 0 Then kind = "CHIEFEVAL"
    If StrComp(Nz(r!ReportType, ""), "FitRep", vbBinaryCompare) = 0 Then kind = "FITREP"
    If kind = "" Then Err.Raise vbObjectError + 414, , "Unknown or null report type is quarantined."
    w.BeginTrans: trans = True
    beforeRevision = Null
    If destinationID > 0 Then
        If CurrentProject.AllForms("frmEditor").IsLoaded Then Err.Raise vbObjectError + 416, , "Close the editor before applying an incoming narrative."
        Set target = db.OpenRecordset("SELECT * FROM Report WHERE ReportID=" & destinationID, dbOpenSnapshot)
        If target.EOF Then Err.Raise vbObjectError + 417, , "Destination draft missing."
        If target!FormKind <> kind Then Err.Raise vbObjectError + 418, , "Source and destination report types differ."
        If target!CurrentRevisionID <> expectedRevision Then Err.Raise vbObjectError + 419, , "Destination changed after preview. Compare again."
        beforeRevision = target!CurrentRevisionID
        id = CommitDraft(destinationID, kind, Nz(target!Title, ""), Nz(target!MemberName, ""), target!PeriodStart, target!PeriodEnd, Nz(target!DraftNotes, ""), "Accept incoming narrative", r!Comments)
        target.Close
    Else
        id = CommitDraft(0, kind, "Imported narrative " & externalID, memberDisplay, Null, Null, "Other official fields require mapping and review. Source values retained in import provenance.", "Import confirmed narrative", r!Comments)
    End If
    r.Close
    Set target = db.OpenRecordset("SELECT CurrentRevisionID FROM Report WHERE ReportID=" & id, dbOpenSnapshot)
    Set receipt = db.OpenRecordset("ImportReceipt", dbOpenDynaset)
    receipt.AddNew: receipt!ReceiptGuid = NewIdentity(): receipt!SourceHash = hash
    receipt!SourcePath = reviewDb.Name: receipt!ExternalReportID = externalID: receipt!ReportID = id
    receipt!AuthorProfileID = DLookup("ActiveProfileID", "Workspace", "WorkspaceID=1")
    receipt!ImportedAt = Now(): receipt!AdapterVersion = "comments-v1/local-CStr"
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
    If SourceFileHash(reviewDb.Name) <> hash Then Err.Raise vbObjectError + 415, , "Source changed during import. Review it again."
    Call FinishOperation(operationID, "Completed", "Applied narrative to Swift report " & id)
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
    sourceHash = SourceFileHash(reviewDb.Name)
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
    If SourceFileHash(reviewDb.Name) <> sourceHash Then Err.Raise vbObjectError + 420, , "Source changed during preview. Choose it again."
    result.Add "SourceHash", sourceHash
    Set NarrativeImportValues = result
End Function

Private Sub ValidateReviewSchema(ByVal db As DAO.Database, ByVal prefix As String)
    Dim t As DAO.TableDef, f As DAO.Field
    Set t = db.TableDefs(prefix & "Folders")
    If t.Attributes <> 0 Then Err.Raise vbObjectError + 404, , "Linked or hidden source tables are unsupported."
    If t.Fields.Count <> 78 Then Err.Raise vbObjectError + 402, , "Unrecognized staging table."
    Set f = t.Fields(0)
    If StrComp(f.Name, "FolderName", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 50 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(1)
    If StrComp(f.Name, "FolderID", vbBinaryCompare) <> 0 Or f.Type <> 4 Or f.Size <> 4 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(2)
    If StrComp(f.Name, "Parent", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(3)
    If StrComp(f.Name, "FullName", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 27 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(4)
    If StrComp(f.Name, "Rate", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 5 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(5)
    If StrComp(f.Name, "Desig", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 12 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(6)
    If StrComp(f.Name, "SSN", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 9 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(7)
    If StrComp(f.Name, "Active", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(8)
    If StrComp(f.Name, "TAR", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(9)
    If StrComp(f.Name, "Inactive", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(10)
    If StrComp(f.Name, "ATADSW", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(11)
    If StrComp(f.Name, "UIC", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 5 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(12)
    If StrComp(f.Name, "ShipStation", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 18 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(13)
    If StrComp(f.Name, "PromotionStatus", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(14)
    If StrComp(f.Name, "DateReported", vbBinaryCompare) <> 0 Or f.Type <> 8 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(15)
    If StrComp(f.Name, "Periodic", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(16)
    If StrComp(f.Name, "DetInd", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(17)
    If StrComp(f.Name, "Frocking", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(18)
    If StrComp(f.Name, "Special", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(19)
    If StrComp(f.Name, "FromDate", vbBinaryCompare) <> 0 Or f.Type <> 8 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(20)
    If StrComp(f.Name, "ToDate", vbBinaryCompare) <> 0 Or f.Type <> 8 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(21)
    If StrComp(f.Name, "NOB", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(22)
    If StrComp(f.Name, "Regular", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(23)
    If StrComp(f.Name, "Concurrent", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(24)
    If StrComp(f.Name, "OpsCdr", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(25)
    If StrComp(f.Name, "PhysicalReadiness", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 4 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(26)
    If StrComp(f.Name, "BilletSubcat", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 10 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(27)
    If StrComp(f.Name, "ReportingSenior", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 18 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(28)
    If StrComp(f.Name, "RSGrade", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 5 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(29)
    If StrComp(f.Name, "RSDesig", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 5 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(30)
    If StrComp(f.Name, "RSTitle", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 14 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(31)
    If StrComp(f.Name, "RSUIC", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 5 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(32)
    If StrComp(f.Name, "RSSSN", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 9 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(33)
    If StrComp(f.Name, "Achievements", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(34)
    If StrComp(f.Name, "PrimaryDuty", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 14 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(35)
    If StrComp(f.Name, "Duties", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(36)
    If StrComp(f.Name, "DateCounseled", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(37)
    If StrComp(f.Name, "Counseler", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 20 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(38)
    If StrComp(f.Name, "PROF", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(39)
    If StrComp(f.Name, "QUAL", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(40)
    If StrComp(f.Name, "EO", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(41)
    If StrComp(f.Name, "MIL", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(42)
    If StrComp(f.Name, "PA", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(43)
    If StrComp(f.Name, "TEAM", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(44)
    If StrComp(f.Name, "LEAD", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(45)
    If StrComp(f.Name, "MIS", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(46)
    If StrComp(f.Name, "TAC", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(47)
    If StrComp(f.Name, "RecommendA", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 20 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(48)
    If StrComp(f.Name, "RecommendB", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 20 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(49)
    If StrComp(f.Name, "Rater", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 18 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(50)
    If StrComp(f.Name, "RaterDate", vbBinaryCompare) <> 0 Or f.Type <> 8 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(51)
    If StrComp(f.Name, "CommentsPitch", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(52)
    If StrComp(f.Name, "Comments", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(53)
    If StrComp(f.Name, "Qualifications", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(54)
    If StrComp(f.Name, "PromotionRecom", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(55)
    If StrComp(f.Name, "SummarySP", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(56)
    If StrComp(f.Name, "SummaryProg", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(57)
    If StrComp(f.Name, "SummaryProm", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(58)
    If StrComp(f.Name, "SummaryMP", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(59)
    If StrComp(f.Name, "SummaryEP", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(60)
    If StrComp(f.Name, "RetentionYes", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(61)
    If StrComp(f.Name, "RetentionNo", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(62)
    If StrComp(f.Name, "RSAddress", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(63)
    If StrComp(f.Name, "RSAddress1", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 30 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(64)
    If StrComp(f.Name, "RSAddress2", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 30 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(65)
    If StrComp(f.Name, "RSCity", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 15 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(66)
    If StrComp(f.Name, "RSState", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(67)
    If StrComp(f.Name, "RSZipCd", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 9 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(68)
    If StrComp(f.Name, "RSPhone", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 10 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(69)
    If StrComp(f.Name, "RSDSN", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 7 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(70)
    If StrComp(f.Name, "SeniorRater", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 18 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(71)
    If StrComp(f.Name, "SeniorRaterDate", vbBinaryCompare) <> 0 Or f.Type <> 8 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(72)
    If StrComp(f.Name, "StatementYes", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(73)
    If StrComp(f.Name, "StatementNo", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(74)
    If StrComp(f.Name, "RSInfo", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(75)
    If StrComp(f.Name, "AutoSummary", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(76)
    If StrComp(f.Name, "UserComments", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(77)
    If StrComp(f.Name, "Psswrd", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set t = db.TableDefs(prefix & "Reports")
    If t.Attributes <> 0 Then Err.Raise vbObjectError + 404, , "Linked or hidden source tables are unsupported."
    If t.Fields.Count <> 128 Then Err.Raise vbObjectError + 402, , "Unrecognized staging table."
    Set f = t.Fields(0)
    If StrComp(f.Name, "Parent", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 12 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(1)
    If StrComp(f.Name, "ReportID", vbBinaryCompare) <> 0 Or f.Type <> 4 Or f.Size <> 4 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(2)
    If StrComp(f.Name, "ReportType", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 15 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(3)
    If StrComp(f.Name, "FullName", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 27 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(4)
    If StrComp(f.Name, "FirstName", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 27 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(5)
    If StrComp(f.Name, "MI", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(6)
    If StrComp(f.Name, "LastName", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 27 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(7)
    If StrComp(f.Name, "Suffix", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 10 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(8)
    If StrComp(f.Name, "Rate", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 5 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(9)
    If StrComp(f.Name, "Desig", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 12 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(10)
    If StrComp(f.Name, "SSN", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 9 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(11)
    If StrComp(f.Name, "DoDID", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 10 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(12)
    If StrComp(f.Name, "Active", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(13)
    If StrComp(f.Name, "TAR", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(14)
    If StrComp(f.Name, "Inactive", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(15)
    If StrComp(f.Name, "ATADSW", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(16)
    If StrComp(f.Name, "UIC", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 5 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(17)
    If StrComp(f.Name, "ShipStation", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 18 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(18)
    If StrComp(f.Name, "PromotionStatus", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(19)
    If StrComp(f.Name, "DateReported", vbBinaryCompare) <> 0 Or f.Type <> 8 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(20)
    If StrComp(f.Name, "Periodic", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(21)
    If StrComp(f.Name, "DetInd", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(22)
    If StrComp(f.Name, "Frocking", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(23)
    If StrComp(f.Name, "Special", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(24)
    If StrComp(f.Name, "FromDate", vbBinaryCompare) <> 0 Or f.Type <> 8 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(25)
    If StrComp(f.Name, "ToDate", vbBinaryCompare) <> 0 Or f.Type <> 8 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(26)
    If StrComp(f.Name, "NOB", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(27)
    If StrComp(f.Name, "Regular", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(28)
    If StrComp(f.Name, "Concurrent", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(29)
    If StrComp(f.Name, "OpsCdr", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(30)
    If StrComp(f.Name, "PhysicalReadiness", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 4 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(31)
    If StrComp(f.Name, "PhysicalReadiness2", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 15 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(32)
    If StrComp(f.Name, "PhysicalReadinessDt", vbBinaryCompare) <> 0 Or f.Type <> 8 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(33)
    If StrComp(f.Name, "BilletSubcat", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 10 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(34)
    If StrComp(f.Name, "RSLastName", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 18 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(35)
    If StrComp(f.Name, "RSFI", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(36)
    If StrComp(f.Name, "RSMI", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(37)
    If StrComp(f.Name, "ReportingSenior", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 18 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(38)
    If StrComp(f.Name, "RSGrade", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 5 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(39)
    If StrComp(f.Name, "RSDesig", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 5 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(40)
    If StrComp(f.Name, "RSTitle", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 14 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(41)
    If StrComp(f.Name, "RSUIC", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 5 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(42)
    If StrComp(f.Name, "RSSSN", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 9 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(43)
    If StrComp(f.Name, "RSDoDID", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 10 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(44)
    If StrComp(f.Name, "Achievements", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(45)
    If StrComp(f.Name, "PrimaryDuty", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 14 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(46)
    If StrComp(f.Name, "Duties", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(47)
    If StrComp(f.Name, "DateCounseled", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(48)
    If StrComp(f.Name, "Counseler", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 20 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(49)
    If StrComp(f.Name, "CounselerLN", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 20 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(50)
    If StrComp(f.Name, "CounselerFI", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(51)
    If StrComp(f.Name, "CounselerMI", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(52)
    If StrComp(f.Name, "PROF", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(53)
    If StrComp(f.Name, "PROFDN1", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(54)
    If StrComp(f.Name, "PROFDN2", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(55)
    If StrComp(f.Name, "PROFDN3", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(56)
    If StrComp(f.Name, "QUAL", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(57)
    If StrComp(f.Name, "QUALDN1", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(58)
    If StrComp(f.Name, "QUALDN2", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(59)
    If StrComp(f.Name, "QUALDN3", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(60)
    If StrComp(f.Name, "EO", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(61)
    If StrComp(f.Name, "EODN1", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(62)
    If StrComp(f.Name, "EODN2", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(63)
    If StrComp(f.Name, "EODN3", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(64)
    If StrComp(f.Name, "MIL", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(65)
    If StrComp(f.Name, "MILDN1", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(66)
    If StrComp(f.Name, "MILDN2", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(67)
    If StrComp(f.Name, "MILDN3", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(68)
    If StrComp(f.Name, "PA", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(69)
    If StrComp(f.Name, "PADN1", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(70)
    If StrComp(f.Name, "PADN2", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(71)
    If StrComp(f.Name, "PADN3", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(72)
    If StrComp(f.Name, "TEAM", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(73)
    If StrComp(f.Name, "TEAMDN1", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(74)
    If StrComp(f.Name, "TEAMDN2", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(75)
    If StrComp(f.Name, "TEAMDN3", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(76)
    If StrComp(f.Name, "LEAD", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(77)
    If StrComp(f.Name, "LEADDN1", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(78)
    If StrComp(f.Name, "LEADDN2", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(79)
    If StrComp(f.Name, "LEADDN3", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(80)
    If StrComp(f.Name, "MIS", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(81)
    If StrComp(f.Name, "MISDN1", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(82)
    If StrComp(f.Name, "MISDN2", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(83)
    If StrComp(f.Name, "MISDN3", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(84)
    If StrComp(f.Name, "TAC", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(85)
    If StrComp(f.Name, "TACDN1", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(86)
    If StrComp(f.Name, "TACDN2", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(87)
    If StrComp(f.Name, "TACDN3", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 255 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(88)
    If StrComp(f.Name, "RecommendA", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 20 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(89)
    If StrComp(f.Name, "RecommendB", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 20 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(90)
    If StrComp(f.Name, "Rater", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 28 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(91)
    If StrComp(f.Name, "RaterDate", vbBinaryCompare) <> 0 Or f.Type <> 8 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(92)
    If StrComp(f.Name, "CommentsPitch", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(93)
    If StrComp(f.Name, "Comments", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(94)
    If StrComp(f.Name, "Qualifications", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(95)
    If StrComp(f.Name, "PromotionRecom", vbBinaryCompare) <> 0 Or f.Type <> 3 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(96)
    If StrComp(f.Name, "SummaryRank", vbBinaryCompare) <> 0 Or f.Type <> 4 Or f.Size <> 4 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(97)
    If StrComp(f.Name, "SummarySP", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(98)
    If StrComp(f.Name, "SummaryProg", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(99)
    If StrComp(f.Name, "SummaryProm", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(100)
    If StrComp(f.Name, "SummaryMP", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(101)
    If StrComp(f.Name, "SummaryEP", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(102)
    If StrComp(f.Name, "RetentionYes", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(103)
    If StrComp(f.Name, "RetentionNo", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(104)
    If StrComp(f.Name, "RSCA", vbBinaryCompare) <> 0 Or f.Type <> 20 Or f.Size <> 16 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(105)
    If StrComp(f.Name, "RSAddress", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(106)
    If StrComp(f.Name, "RSAddress1", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 30 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(107)
    If StrComp(f.Name, "RSAddress2", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 30 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(108)
    If StrComp(f.Name, "RSCity", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 15 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(109)
    If StrComp(f.Name, "RSState", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 2 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(110)
    If StrComp(f.Name, "RSZipCd", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 9 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(111)
    If StrComp(f.Name, "RSPhone", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 10 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(112)
    If StrComp(f.Name, "RSDSN", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 7 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(113)
    If StrComp(f.Name, "SeniorRater", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 28 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(114)
    If StrComp(f.Name, "SeniorRaterDate", vbBinaryCompare) <> 0 Or f.Type <> 8 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(115)
    If StrComp(f.Name, "StatementYes", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(116)
    If StrComp(f.Name, "StatementNo", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(117)
    If StrComp(f.Name, "RSInfo", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(118)
    If StrComp(f.Name, "RRSFI", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(119)
    If StrComp(f.Name, "RRSMI", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 3 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(120)
    If StrComp(f.Name, "RRSLastName", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 25 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(121)
    If StrComp(f.Name, "RRSGrade", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 5 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(122)
    If StrComp(f.Name, "RRSCommand", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 30 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(123)
    If StrComp(f.Name, "RRSUIC", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 5 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(124)
    If StrComp(f.Name, "UserComments", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(125)
    If StrComp(f.Name, "Psswrd", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 8 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(126)
    If StrComp(f.Name, "Standards", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 30 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(127)
    If StrComp(f.Name, "IsValidated", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set t = db.TableDefs(prefix & "Summary")
    If t.Attributes <> 0 Then Err.Raise vbObjectError + 404, , "Linked or hidden source tables are unsupported."
    If t.Fields.Count <> 6 Then Err.Raise vbObjectError + 402, , "Unrecognized staging table."
    Set f = t.Fields(0)
    If StrComp(f.Name, "ProfileName", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 50 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(1)
    If StrComp(f.Name, "OnTimeYes", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(2)
    If StrComp(f.Name, "OnTimeNo", vbBinaryCompare) <> 0 Or f.Type <> 1 Or f.Size <> 1 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(3)
    If StrComp(f.Name, "ISICAddress", vbBinaryCompare) <> 0 Or f.Type <> 12 Or f.Size <> 0 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(4)
    If StrComp(f.Name, "EMail", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 50 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
    Set f = t.Fields(5)
    If StrComp(f.Name, "Phone", vbBinaryCompare) <> 0 Or f.Type <> 10 Or f.Size <> 10 Then Err.Raise vbObjectError + 403, , "Unrecognized staging field."
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
        !txtSourceHash = SourceFileHash(reviewDb.Name)
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
    If SourceFileHash(reviewDb.Name) <> expectedHash Then Err.Raise vbObjectError + 422, , "Source changed after preview. Compare again."
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
