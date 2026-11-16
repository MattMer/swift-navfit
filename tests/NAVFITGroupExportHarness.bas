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
Private Function ISICPreviewRejected(ByVal groupID As Long, ByVal isicID As Long) As String
    On Error GoTo Failed
    Call NAVFITGroupPreview(groupID, False, isicID)
    ISICPreviewRejected = "ACCEPTED": Exit Function
Failed:
    ISICPreviewRejected = Err.Description
End Function
Private Function ISICExportRejected(ByVal groupID As Long, ByVal file As String, ByVal template As String, ByVal token As String, ByVal isicID As Long) As String
    On Error GoTo Failed
    Call NAVFITExportGroup(groupID, file, template, token, False, isicID)
    ISICExportRejected = "ACCEPTED": Exit Function
Failed:
    ISICExportRejected = Err.Description
End Function
Private Function SummaryProfiles(ByVal file As String) As Long
    Dim db As DAO.Database, r As DAO.Recordset
    Set db = DBEngine.OpenDatabase(file, False, True)
    Set r = db.OpenRecordset("SELECT Count(*) AS N FROM Summary", dbOpenSnapshot)
    SummaryProfiles = r!N: r.Close: db.Close
End Function
'The optional ISIC becomes a NAVFIT98A summary letter profile: the ISIC address, the
'reporting senior's e-mail and ten-digit phone, on-time unanswered.
Public Function TestNAVFITISICProfile(ByVal directory As String, ByVal template As String) As String
    Dim values As Object, senior As Object, key As Variant, copy As Long, groupID As Long, seniorID As Long, baselineID As Long, isicID As Long, kind As String, grade As String
    Dim preview As Object, token As String, plainToken As String, file As String, native As DAO.Database, r As DAO.Recordset, address As String, blankProfiles As Long
    On Error GoTo Failed
    Set values = SavedFields(1): kind = values("FormKind")
    Select Case kind
        Case "EVAL": grade = "E5"
        Case "CHIEFEVAL": grade = "E7"
        Case "FITREP": grade = "O3"
    End Select
    copy = DuplicateDraft(1): values("ReportingSeniorPhone") = "5551234567"
    Call NAVFITCommitFields(copy, values, "Synthetic ISIC member", "ISIC profile test")
    groupID = CreateSummaryGroup("Synthetic ISIC group", kind, grade): AddGroupMember groupID, copy, grade
    Set senior = RSSeniorValues(0)
    For Each key In RSSeniorKeys(): senior(CStr(key)) = values(CStr(key)): Next
    senior("ReportingSeniorEmail") = "RS.SYNTHETIC@EXAMPLE.MIL"
    seniorID = RSSaveSenior(0, "Synthetic ISIC senior", senior)
    baselineID = RSAddBaseline(seniorID, grade, DateSerial(2026, 1, 1), 4#, 7, 28#, "Synthetic ISIC history")
    RSSetGroupBaseline groupID, seniorID, baselineID, True
    address = "COMMANDER, SYNTHETIC GROUP TWO" & vbCrLf & "123 EXAMPLE STREET" & vbCrLf & "NORFOLK, VA 23511-0001"
    isicID = ISICSave(0, 0, "Synthetic ISIC", "N12345", address)
    blankProfiles = SummaryProfiles(template)
    'No ISIC: the file keeps the template's profiles only.
    Set preview = NAVFITGroupPreview(groupID): plainToken = preview("Token")
    file = directory & "\isic-none.accdb"
    Check NAVFITExportGroup(groupID, file, template, plainToken) = file, "Export without ISIC failed"
    Check SummaryProfiles(file) = blankProfiles, "Export without ISIC added a profile"
    'With an ISIC: one profile, as NAVFIT98A saves it.
    Set preview = NAVFITGroupPreview(groupID, False, isicID): token = preview("Token")
    Check token <> plainToken, "ISIC choice is not part of the preview token"
    Check InStr(preview("Text"), "Synthetic ISIC") > 0, "Preview does not name the profile"
    Check InStr(ISICExportRejected(groupID, directory & "\isic-mismatch.accdb", template, plainToken, isicID), "Preview again") > 0, "Export with an ISIC accepted a preview without it"
    file = directory & "\isic-profile.accdb"
    Check NAVFITExportGroup(groupID, file, template, token, False, isicID) = file, "Export with ISIC failed"
    Set native = DBEngine.OpenDatabase(file, False, True): ValidateNAVFITDatabase native: ValidateNAVFITSourceSchema native, ""
    Set r = native.OpenRecordset("SELECT * FROM Summary WHERE ProfileName='Synthetic ISIC'", dbOpenSnapshot)
    Check Not r.EOF, "Profile missing"
    Check r!ISICAddress = address, "ISIC address differs"
    Check r!EMail = "RS.SYNTHETIC@EXAMPLE.MIL", "E-mail differs"
    Check r!Phone = "5551234567", "Phone differs"
    Check Not r!OnTimeYes And Not r!OnTimeNo, "On time was answered"
    r.Close: native.Close
    Check SummaryProfiles(file) = blankProfiles + 1, "Profile count differs"
    'A changed ISIC needs a new preview.
    ISICSave isicID, CLng(DLookup("VersionCounter", "ISIC", "ISICID=" & isicID)), "Synthetic ISIC", "N12345", address & vbCrLf & "ATTN: N1"
    Check InStr(ISICExportRejected(groupID, directory & "\isic-stale.accdb", template, token, isicID), "Preview again") > 0, "Stale ISIC exported"
    'A phone NAVFIT98A cannot hold is left out, with a note.
    values("ReportingSeniorPhone") = "555-1234": Call NAVFITCommitFields(copy, values, "Synthetic ISIC member", "Short phone")
    senior("ReportingSeniorPhone") = "555-1234": Call RSSaveSenior(seniorID, "Synthetic ISIC senior", senior, CLng(DLookup("VersionCounter", "ReportingSeniorRecord", "SeniorID=" & seniorID)))
    Set preview = NAVFITGroupPreview(groupID, False, isicID)
    Check preview("ISICProfile")("Phone") = "" And InStr(preview("ISICProfile")("Note"), "ten-digit") > 0, "Short phone not reported"
    file = directory & "\isic-short-phone.accdb"
    Check NAVFITExportGroup(groupID, file, template, preview("Token"), False, isicID) = file, "Export with short phone failed"
    Set native = DBEngine.OpenDatabase(file, False, True)
    Set r = native.OpenRecordset("SELECT * FROM Summary WHERE ProfileName='Synthetic ISIC'", dbOpenSnapshot)
    Check IsNull(r!Phone), "Short phone written": r.Close: native.Close
    'A deleted ISIC cannot be chosen.
    ISICDelete isicID
    Check InStr(ISICPreviewRejected(groupID, isicID), "Choose an ISIC") > 0, "Deleted ISIC accepted"
    TestNAVFITISICProfile = "PASS: " & checks & " ISIC profile assertions": Exit Function
Failed:
    TestNAVFITISICProfile = "FAIL: " & Err.Number & " " & Err.Description
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
            Set senior = RSSeniorReportValues(0)
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

'Narrative pasted from Word: curly quotes, dashes, ellipsis, bullet, non-breaking
'space, paragraph and soft line breaks. NAVFIT98A validates characters only in
'blocks 2 and 3, so the export must carry the saved text exactly.
Public Function TestNAVFITPastedText(ByVal directory As String, ByVal template As String) As String
    Dim pasted As String, copy As Long, values As Object, native As DAO.Database, r As DAO.Recordset, file As String, exported As Variant, i As Long
    On Error GoTo Failed
    checks = 0
    pasted = "Led " & ChrW(8220) & "Bravo" & ChrW(8221) & " section" & ChrW(8217) & "s 24" & ChrW(8211) & "hour watch" & ChrW(8212) & "zero mishaps" & ChrW(8230) & ChrW(160) & ChrW(8226) & " Qualified EOOW." & vbCrLf & "Second paragraph" & Chr$(11) & "after a soft break."
    copy = DuplicateDraft(1)
    Set values = SavedFields(copy)
    values("OfficialComments") = pasted
    Call NAVFITCommitFields(copy, values, "Synthetic pasted text", "Pasted text test")
    Check StrComp(Nz(SavedFields(copy)("OfficialComments"), ""), pasted, vbBinaryCompare) = 0, "Saving changed pasted narrative text"
    file = directory & "\pasted-text.accdb"
    Check NAVFITExportReport(copy, file, template) = file, "Pasted narrative blocked export"
    Set native = DBEngine.OpenDatabase(file, False, True)
    Set r = native.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
    Check Not r.EOF, "Exported report missing"
    exported = NAVFITReadFields(r, 2000)("OfficialComments")
    r.Close: native.Close: Set native = Nothing
    If StrComp(Nz(exported, ""), pasted, vbBinaryCompare) <> 0 Then
        For i = 1 To Len(pasted)
            If Mid$(Nz(exported, ""), i, 1) <> Mid$(pasted, i, 1) Then Exit For
        Next
        Check False, "Export changed pasted text at character " & i & ": saved U+" & Hex$(AscW(Mid$(pasted, i, 1) & " ")) & ", exported U+" & Hex$(AscW(Mid$(Nz(exported, "") & " ", i, 1)))
    End If
    checks = checks + 1
    TestNAVFITPastedText = "PASS: " & checks & " Word-pasted narrative save/export assertions": Exit Function
Failed:
    If Not native Is Nothing Then native.Close
    TestNAVFITPastedText = "FAIL: " & Err.Number & " " & Err.Description
End Function

'An unanswered check box does not stop the export: NAVFIT98A boxes are ticked or not,
'so it goes out unticked and the export result names it.
Public Function TestNAVFITBlankBoxes(ByVal directory As String, ByVal template As String) As String
    Dim copy As Long, values As Object, native As DAO.Database, r As DAO.Recordset, file As String
    On Error GoTo Failed
    checks = 0
    copy = DuplicateDraft(1)
    Set values = SavedFields(copy)
    values("StatementYes") = Null: values("StatementNo") = Null
    Call NAVFITCommitFields(copy, values, "Synthetic blank statement", "Blank check box test")
    Check IsNull(SavedFields(copy)("StatementYes")), "Saving filled the blank statement box"
    file = directory & "\blank-boxes.accdb"
    Check NAVFITExportReport(copy, file, template) = file, "A blank check box blocked export"
    Check InStr(NAVFITBlankBoxText(), UXFieldLabel("StatementYes")) > 0 And InStr(NAVFITBlankBoxText(), UXFieldLabel("StatementNo")) > 0, "Blank check boxes not listed in the export result"
    Set native = DBEngine.OpenDatabase(file, False, True)
    Set r = native.OpenRecordset("SELECT * FROM Reports WHERE ReportType Is Not Null", dbOpenSnapshot)
    Check Not r.EOF, "Exported report missing"
    Check NAVFITReadFields(r, 2000)("StatementYes") = "No", "Blank statement box was not exported unticked"
    r.Close: native.Close: Set native = Nothing
    Check IsNull(SavedFields(copy)("StatementYes")), "Export changed the saved blank box"
    TestNAVFITBlankBoxes = "PASS: " & checks & " blank check box export assertions": Exit Function
Failed:
    If Not native Is Nothing Then native.Close
    TestNAVFITBlankBoxes = "FAIL: " & Err.Number & " " & Err.Description
End Function
