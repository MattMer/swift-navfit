Attribute VB_Name = "NAVFITGroupExportHarness"
Option Compare Binary
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    checks = checks + 1
    If Not ok Then Err.Raise vbObjectError + 5360, , message
End Sub
Private Function Rejected(ByVal groupID As Long, ByVal file As String, ByVal template As String, ByVal token As String, Optional ByVal usePSG As Boolean = False) As String
    On Error GoTo Failed
    Call NAVFITExportGroup(groupID, file, template, token, usePSG)
    Rejected = "ACCEPTED": Exit Function
Failed:
    Rejected = Err.Description
End Function
Private Function MemberRejected(ByVal groupID As Long, ByVal reportID As Long) As String
    On Error GoTo Failed
    AddGroupMember groupID, reportID, "WRONG"
    MemberRejected = "ACCEPTED": Exit Function
Failed:
    MemberRejected = Err.Description
End Function
Private Function SummaryRejected(ByVal db As DAO.Database, ByVal preview As Object) As Boolean
    On Error GoTo Rejected
    VerifyNativeGroupSummary db, preview: Exit Function
Rejected:
    SummaryRejected = True
End Function
Public Function TestNAVFITGroupExport(ByVal directory As String, ByVal template As String) As String
    Dim id As Long, copy As Long, groupID As Long, values As Object, preview As Object, native As DAO.Database, baseline As DAO.Database, r As DAO.Recordset, decoded As Object, key As Variant, file As String, kind As String, grade As String, old As Variant, token As String, revisionCount As Long, original As DAO.Database, before As DAO.Recordset, senior As Object, seniorID As Long, baselineID As Long, psg As Object
    On Error GoTo Failed
    For id = 1 To 3
        Set values = SavedFields(id): kind = values("FormKind")
        Select Case kind
            Case "EVAL": grade = "E5"
            Case "CHIEFEVAL": grade = "E7"
            Case "FITREP": grade = "O3"
        End Select
        copy = DuplicateDraft(id)
        values("PromotionRecommendation") = NAVFITDecode(5, "promotion")
        Call NAVFITCommitFields(copy, values, "Synthetic group member", "Group export test")
        groupID = CreateSummaryGroup("Synthetic " & kind & " group", kind, grade)
        Check InStr(MemberRejected(groupID, id), "paygrade must match") > 0, "Membership mismatch masked by cleanup error"
        AddGroupMember groupID, id, grade: AddGroupMember groupID, copy, grade
        Set preview = NAVFITGroupPreview(groupID): token = preview("Token")
        Check preview("Count5") >= 1, "EP count missing"
        revisionCount = DCount("*", "ReportRevision")
        file = directory & "\group-" & kind & ".accdb"
        Check NAVFITExportGroup(groupID, file, template, token) = file, "Group export failed"
        Check DCount("*", "ReportRevision") = revisionCount, "Export created revisions"
        Set native = DBEngine.OpenDatabase(file, False, True): ValidateNAVFITDatabase native
        Set r = native.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null ORDER BY ReportID", dbOpenSnapshot)
        Check Not r.EOF, "No exported reports"
        Set values = SavedFields(id): Set decoded = NAVFITReadFields(r, 2000)
        For Each key In decoded.Keys: Check SameFieldValue(decoded(key), values(key)), "Mapped group field differs: " & key: Next
        Check r!Parent = "a 1", "Native folder linkage differs"
        Check r!SummaryEP = CStr(preview("Count5")), "Native EP total differs"
        Check Not r!IsValidated, "Export claimed native validation"
        Set original = DBEngine.OpenDatabase(CStr(DLookup("SourcePath", "ImportReceipt", "ReportID=" & id)), False, True)
        Set before = original.OpenRecordset("SELECT * FROM Reports WHERE ReportID=" & DLookup("ExternalReportID", "ImportReceipt", "ReportID=" & id), dbOpenSnapshot)
        For Each key In before.Fields
            Select Case key.Name
                Case "ReportID", "Parent", "SummarySP", "SummaryProg", "SummaryProm", "SummaryMP", "SummaryEP", "IsValidated"
                Case Else: Check SameFieldValue(key.Value, r.Fields(key.Name).Value), "Native-only preservation failed: " & key.Name
            End Select
        Next
        before.Close: original.Close
        r.MoveNext: Check Not r.EOF, "Second member missing"
        Set values = SavedFields(copy): Set decoded = NAVFITReadFields(r, 2000)
        For Each key In decoded.Keys: Check SameFieldValue(decoded(key), values(key)), "New draft field differs: " & key: Next
        r.MoveNext: Check r.EOF, "Unselected report exported"
        r.Close
        Set r = native.OpenRecordset("SELECT * FROM Folders", dbOpenSnapshot)
        Check r!FolderName = "Synthetic " & kind & " group", "Folder title differs"
        Check r!AutoSummary, "Native auto-summary disabled"
        r.MoveNext: Check r.EOF, "Unselected folder exported": r.Close: native.Close
        Set baseline = DBEngine.OpenDatabase(file & ".swift-baseline.accdb", False, True)
        Set r = baseline.OpenRecordset("SELECT Count(*) AS N FROM SwiftGroupReport", dbOpenSnapshot)
        Check r!N = 2, "Baseline member identities missing": r.Close
        Set r = baseline.OpenRecordset("SELECT Count(*) AS N FROM SwiftGroupValue", dbOpenSnapshot)
        Check r!N = 2 * (UBound(DraftFieldKeys()) + 1), "Baseline does not freeze all current saved fields": r.Close: baseline.Close
        FileCopy file, file & ".corrupt.accdb": Set native = DBEngine.OpenDatabase(file & ".corrupt.accdb")
        native.Execute "UPDATE Reports SET SummaryEP='999' WHERE ReportType Is Not Null", dbFailOnError
        Check SummaryRejected(native, preview), "Corrupted native summary was accepted": native.Close
        Check InStr(Rejected(groupID, file, template, token), "not overwritten") > 0, "Existing export overwritten"
        Set values = SavedFields(copy): old = values("ReportingSeniorSSN"): values("ReportingSeniorSSN") = "999999999"
        Call NAVFITCommitFields(copy, values, "Synthetic group member", "Mismatch test")
        Set preview = NAVFITGroupPreview(groupID)
        Check NAVFITExportGroup(groupID, directory & "\mixed-" & id & ".accdb", template, preview("Token")) <> "", "Shared-field findings blocked export"
        values("ReportingSeniorSSN") = old: Call NAVFITCommitFields(copy, values, "Synthetic group member", "Restore test")
        Check InStr(Rejected(groupID, directory & "\stale-" & id & ".accdb", template, token), "Preview again") > 0, "Stale preview exported"
        If id = 1 Then
            Set senior = RSSeniorValues(0)
            For Each key In senior.Keys: senior(key) = values(key): Next
            seniorID = RSSaveSenior(0, "Synthetic export senior", senior)
            baselineID = RSAddBaseline(seniorID, grade, DateSerial(2026, 1, 1), 4#, 7, 28#, "Synthetic exact history")
            RSSetGroupBaseline groupID, seniorID, baselineID, True
            Set preview = NAVFITGroupPreview(groupID, True): Set psg = RSGroupPSG(groupID)
            Check NAVFITExportGroup(groupID, directory & "\group-exact-psg.accdb", template, preview("Token"), True) <> "", "Exact PSG export failed"
            Set native = DBEngine.OpenDatabase(directory & "\group-exact-psg.accdb", False, True)
            Set r = native.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
            Do While Not r.EOF
                Check Abs(CDbl(r!RSCA) - Round(CDbl(psg("Value")), 2)) < 0.00001, "Exact PSG not written"
                r.MoveNext
            Loop
            r.Close: native.Close
        End If
    Next
    TestNAVFITGroupExport = "PASS: " & checks & " group export assertions": Exit Function
Failed:
    TestNAVFITGroupExport = "FAIL: " & Err.Number & " " & Err.Description
End Function
Public Function TestNAVFITNonblocking(ByVal directory As String, ByVal template As String) As String
    Dim id As Long, member As Long, groupID As Long, values As Object, preview As Object, r As DAO.Recordset, native As DAO.Database, baseline As DAO.Database
    Dim file As String, revisions As Long, receipts As Long, raw As Long, stamp As String, recovery As Long, findings As Collection, before As Object, key As Variant
    On Error GoTo Failed
    groupID = CreateSummaryGroup("Synthetic incomplete mixed group", "EVAL", "E5")
    For id = 1 To 3
        Set values = SavedFields(id)
        values("ReportingSeniorSSN") = Null: values("PeriodEnd") = Null: values("PromotionRecommendation") = Null
        values("Trait33") = Null: values("OfficialComments") = Null
        member = NAVFITCommitFields(0, values, "Synthetic incomplete " & id, "Nonblocking export test")
        ' Deliberately model an existing invalid mixed roster without weakening membership writes.
        CurrentDb.Execute "INSERT INTO GroupMember (GroupID,ReportID,Paygrade) VALUES (" & groupID & "," & member & ",'E5')", dbFailOnError
        CurrentDb.Execute "UPDATE Report SET IsArchived=True,WorkflowState='Approved' WHERE ReportID=" & member, dbFailOnError
    Next
    CurrentDb.Execute "UPDATE SummaryGroup SET IsArchived=True WHERE GroupID=" & groupID, dbFailOnError
    Set r = CurrentDb.OpenRecordset("RecoveryDraft", dbOpenDynaset)
    r.AddNew: r!SessionGuid = NewIdentity(): r!AuthorProfileID = ActiveContributorID(): r!ReportID = member: r!BaseRevisionID = CurrentRevision(member)
    r!CheckpointAt = #1/1/2000#: r!DraftTitle = "SYNTHETIC UNSAVED": r.Update: r.Bookmark = r.LastModified: recovery = r!RecoveryDraftID: r.Close
    revisions = DCount("*", "ReportRevision"): receipts = DCount("*", "ImportReceipt"): raw = DCount("*", "ImportRawValue"): stamp = GroupStamp(groupID)
    Set before = SavedFields(member)
    file = directory & "\incomplete-report.accdb"
    Check NAVFITExportReport(member, file, template) = file, "Archived approved incomplete report blocked"
    Set findings = NAVFITLastReportReadiness(): Check findings.Count > 0, "Report findings not available"
    Set preview = NAVFITGroupPreview(groupID, True)
    Set findings = preview("Readiness"): Check findings.Count > 0, "Mixed-group findings missing"
    Check Not preview("WritePSG") And InStr(preview("Text"), "PSG RSCA unavailable") > 0, "Unavailable PSG fallback not disclosed"
    Check preview("Count0") = 0, "Missing recommendations invented NOB counts"
    file = directory & "\incomplete-group.accdb"
    Check NAVFITExportGroup(groupID, file, template, preview("Token"), True) = file, "Incomplete mixed archived group blocked"
    Set native = DBEngine.OpenDatabase(file, False, True)
    Set r = native.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
    id = 0
    Do While Not r.EOF
        id = id + 1
        Check IsNull(r!RSSSN) And IsNull(r!ToDate), "Missing report values invented"
        Set values = NAVFITReadFields(r, 2000)
        Check IsNull(values("PromotionRecommendation")) And IsNull(values("Trait33")), "Missing recommendation/trait invented"
        Check CDbl(r!RSCA) = 0, "Unavailable PSG replaced the native new-row default"
        Check Not r!IsValidated, "Export claimed validation"
        r.MoveNext
    Loop
    Check id = 3, "Mixed/incomplete reports omitted": r.Close: native.Close
    Set baseline = DBEngine.OpenDatabase(file & ".swift-baseline.accdb", False, True)
    Set r = baseline.OpenRecordset("SELECT * FROM SwiftGroupReport", dbOpenSnapshot): r.MoveLast
    Check r.RecordCount = 3, "Baseline omitted mixed reports": r.Close
    Set r = baseline.OpenRecordset("SwiftGroupExport", dbOpenSnapshot): Check Not r!UsePSG, "Baseline claimed unavailable PSG was written": r.Close: baseline.Close
    Check GroupStamp(groupID) = stamp And DCount("*", "ReportRevision") = revisions, "Export changed saved revisions/status/roster"
    Check DCount("*", "ImportReceipt") = receipts And DCount("*", "ImportRawValue") = raw, "Export changed import provenance"
    For Each key In before.Keys: Check SameFieldValue(before(key), SavedFields(member)(key)), "Export changed saved " & key: Next
    Check DLookup("DraftTitle", "RecoveryDraft", "RecoveryDraftID=" & recovery) = "SYNTHETIC UNSAVED", "Export changed pending recovery"
    Check DLookup("CheckpointAt", "RecoveryDraft", "RecoveryDraftID=" & recovery) = #1/1/2000#, "Export rewrote recovery checkpoint"
    Check InStr(Rejected(groupID, file, template, preview("Token"), True), "not overwritten") > 0, "Collision accepted"
    Check Rejected(groupID, directory & "\missing-folder\export.accdb", template, preview("Token"), True) <> "ACCEPTED", "Inaccessible destination accepted"
    CurrentDb.Execute "DELETE FROM RecoveryDraft WHERE RecoveryDraftID=" & recovery, dbFailOnError
    TestNAVFITNonblocking = "PASS: " & checks & " cumulative group/nonblocking export assertions": Exit Function
Failed:
    TestNAVFITNonblocking = "FAIL: " & Err.Number & " " & Err.Source & " " & Err.Description
End Function
