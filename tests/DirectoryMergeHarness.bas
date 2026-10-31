Attribute VB_Name = "DirectoryMergeHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 9190, "DirectoryMergeHarness", message
    checks = checks + 1
End Sub
Public Function TestMergeDisplayFailure() As String
    Dim first As Long, second As Long, merged As Long
    On Error GoTo Failed
    checks = 0: ConfigureWorkspace "SYNTHETIC MERGE DISPLAY FAILURE", "Synthetic merger"
    first = SaveDirectoryMember(0, 0, "Synthetic Original", "ORIGINAL, SYNTHETIC", "IT2", "")
    second = SaveDirectoryMember(0, 0, "Synthetic Survivor", "SURVIVOR, SYNTHETIC", "IT2", "")
    DoCmd.OpenForm "frmUXDirectoryMerge"
    Forms!frmUXDirectoryMerge!cboSource = first: Forms!frmUXDirectoryMerge!cboTarget = second
    Forms!frmUXDirectoryMerge!txtReason = "Synthetic duplicate reviewed"
    Check UXDirectoryMergePreview(), "Display-fault merge preview failed"
    Check Not UXDirectoryMergeApply(), "Injected merge display failure was accepted"
    Check DCount("*", "DirectoryMerge") = 1 And DLookup("IsArchived", "CommandMember", "MemberID=" & first), "Display failure lost the committed merge"
    Check InStr(Forms!frmUXDirectoryMerge!txtPreview, "committed") > 0 And InStr(Forms!frmUXDirectoryMerge!txtPreview, "SYNTHETIC") > 0, "Committed merge/display failure not disclosed"
    Check Not Forms!frmUXDirectoryMerge!cmdApply.Enabled And Not UXDirectoryMergeApply(), "Display failure retained an actionable merge"
    Check DCount("*", "DirectoryMerge") = 1, "Repeated Apply duplicated merge history"
    merged = DMax("MergeID", "DirectoryMerge")
    Check UndoDirectoryMerge(merged), "Display failure broke guarded undo"
    Check Not DLookup("IsArchived", "CommandMember", "MemberID=" & first) And DLookup("MergeState", "DirectoryMerge", "MergeID=" & merged) = "Undone", "Undo lost member availability or merge history"
    TestMergeDisplayFailure = "PASS: " & checks & " committed-merge/display-failure/undo assertions": Exit Function
Failed:
    TestMergeDisplayFailure = "FAIL: merge display fault: " & Err.Number & " " & Err.Description
End Function
Private Function RejectPreview(ByVal a As Long, ByVal b As Long) As Boolean
    On Error GoTo Rejected
    Dim stamp As String: stamp = PreviewDirectoryMerge(a, b): Exit Function
Rejected:
    RejectPreview = True
End Function
Private Function RejectApply(ByVal a As Long, ByVal b As Long, ByVal stamp As String, Optional ByVal reason As String = "Reviewed same Sailor") As Boolean
    On Error GoTo Rejected
    Dim id As Long: id = ApplyDirectoryMerge(a, b, stamp, reason): Exit Function
Rejected:
    RejectApply = True
End Function
Private Function RejectUndo(ByVal id As Long) As Boolean
    On Error GoTo Rejected
    UndoDirectoryMerge id: Exit Function
Rejected:
    RejectUndo = True
End Function
Public Function TestDirectoryMerge() As String
    Dim a As Long, b As Long, c As Long, report As Long, prior As Long, merged As Long, rev As Long, count As Long
    Dim beforeArchive As Variant, stamp As String, before As String, db As DAO.Database, marks As Object, block As Long, opCount As Long, field As Variant
    Dim originalLink As DAO.Recordset, originalPrior As DAO.Recordset, savedLink As Object, savedPrior As Object
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC DIRECTORY MERGE", "Merge tester"
    a = SaveDirectoryMember(0, 0, "Casey Morgan", "MORGAN, CASEY", "IT2", "")
    b = SaveDirectoryMember(0, 0, "Casey Morgan", "MORGAN, CASEY", "IT2", "000000005")
    c = SaveDirectoryMember(0, 0, "Taylor Morgan", "MORGAN, TAYLOR", "IT2", "000000006")
    Check a <> b And b <> c, "Surname auto-merged people"
    Check DirectorySurname("Casey O'Neil", "O'NEIL, CASEY") = DirectorySurname("Casey Oneil", "ONEIL, CASEY"), "Surname normalization"
    Check LongestDigitRun("ab1234567890") = 10 And LongestDigitRun("12-345") = 3 And LongestDigitRun("") = 0, "Digit run length"
    Check RejectPreview(b, c), "Conflicting profile SSNs accepted"
    Check RejectPreview(b, a), "SSN-owning source lost survivor preference"
    Check RejectPreview(a, a) And RejectPreview(a, 999999), "Invalid pair accepted"
    DirectoryCreationContext a
    report = CommitDraft(0, "EVAL", "Merge preservation", "Casey Morgan", #1/1/2026#, #9/30/2026#, "", "Test", "")
    DirectoryCreationContext 0
    rev = CurrentRevision(report): count = DCount("*", "ReportRevision")
    Set db = CurrentDb
    db.Execute "UPDATE MemberReport SET MatchState='Provisional',MatchSource='Original test link' WHERE ReportID=" & report, dbFailOnError
    Set marks = CreateObject("Scripting.Dictionary")
    For block = 33 To 39: marks.Add "Trait" & block, "4.0": Next
    prior = SavePriorEvaluation(0, 0, a, "EVAL", "IT3", #1/1/2025#, #12/31/2025#, "Promotable", "Observed", marks, "Synthetic source")
    Set savedLink = CreateObject("Scripting.Dictionary"): Set savedPrior = CreateObject("Scripting.Dictionary")
    Set originalLink = db.OpenRecordset("SELECT * FROM MemberReport WHERE ReportID=" & report, dbOpenSnapshot)
    For Each field In Array("MatchState", "MatchSource", "UpdatedBy", "UpdatedAt"): savedLink.Add CStr(field), originalLink.Fields(field).Value: Next
    originalLink.Close
    Set originalPrior = db.OpenRecordset("SELECT * FROM PriorEvaluation WHERE PriorID=" & prior, dbOpenSnapshot)
    For Each field In Array("ModifiedBy", "ModifiedAt"): savedPrior.Add CStr(field), originalPrior.Fields(field).Value: Next
    originalPrior.Close
    stamp = PreviewDirectoryMerge(a, b)
    db.Execute "UPDATE CommandMember SET FullName='Casey Newname',ReportName='NEWNAME, CASEY',VersionCounter=VersionCounter+1 WHERE MemberID=" & a, dbFailOnError
    Check RejectApply(a, b, stamp), "Stale profile preview applied"
    stamp = PreviewDirectoryMerge(a, b)
    db.Execute "UPDATE MemberReport SET MatchSource='Changed link' WHERE ReportID=" & report, dbFailOnError
    Check RejectApply(a, b, stamp), "Stale link preview applied"
    db.Execute "UPDATE MemberReport SET MatchSource='Original test link' WHERE ReportID=" & report, dbFailOnError
    db.Execute "UPDATE Report SET MemberDoDID='0000000001' WHERE ReportID=" & report, dbFailOnError
    DirectoryCreationContext b
    Dim otherReport As Long
    otherReport = CommitDraft(0, "EVAL", "Other linked record", "Casey Morgan", #1/1/2025#, #9/30/2025#, "", "Test", "")
    DirectoryCreationContext 0
    db.Execute "UPDATE Report SET MemberDoDID='0000000002' WHERE ReportID=" & otherReport, dbFailOnError
    Check RejectPreview(a, b), "Conflicting report DoD IDs accepted"
    'Restore both injected cache values: a cache that disagrees with the saved
    'revision is (correctly) refused by workspace import in the release gate.
    db.Execute "UPDATE Report SET MemberDoDID=Null WHERE ReportID IN (" & report & "," & otherReport & ")", dbFailOnError
    db.Execute "UPDATE Report SET MemberSSN='000000006' WHERE ReportID=" & report, dbFailOnError
    Check RejectPreview(a, b), "Conflicting linked report SSN accepted"
    db.Execute "UPDATE Report SET MemberSSN=Null WHERE ReportID=" & report, dbFailOnError
    stamp = PreviewDirectoryMerge(a, b)
    Check RejectApply(a, b, stamp, ""), "Blank reason accepted"
    Check RejectApply(a, b, stamp, "SSN 000-00-0005"), "Identifier retained in merge reason"
    before = DirectoryMergeStamp(a, b): opCount = DCount("*", "OperationJournal")
    db.TableDefs("CommandMember").ValidationRule = "[IsArchived]=False"
    Check RejectApply(a, b, stamp), "Injected archive failure did not reject"
    db.TableDefs("CommandMember").ValidationRule = ""
    DBEngine.Idle dbRefreshCache
    Check DirectoryMergeStamp(a, b) = before, "Partial merge was not rolled back: " & before & " -> " & DirectoryMergeStamp(a, b)
    Check DCount("*", "DirectoryMerge") = 0 And DCount("*", "DirectoryMergeReport") = 0 And DCount("*", "DirectoryMergePrior") = 0 And DCount("*", "OperationJournal") = opCount, "Rollback leaked audit/mapping rows"
    merged = ApplyDirectoryMerge(a, b, stamp, "Reviewed documents confirm name change")
    Check DLookup("MemberID", "MemberReport", "ReportID=" & report) = b, "Report link not moved"
    Check DLookup("MemberID", "PriorEvaluation", "PriorID=" & prior) = b, "Prior link not moved"
    Check DLookup("IsArchived", "CommandMember", "MemberID=" & a), "Original profile not archived"
    Check DLookup("FullName", "CommandMember", "MemberID=" & a) = "Casey Newname", "Original name lost"
    Check DLookup("MemberSSN", "CommandMember", "MemberID=" & b) = "000000005", "Survivor identity changed"
    Check DLookup("MatchState", "MemberReport", "ReportID=" & report) = "Provisional", "Unconfirmed link was silently confirmed"
    Check CurrentRevision(report) = rev And DLookup("MemberName", "Report", "ReportID=" & report) = "Casey Morgan", "Report/revision changed"
    Check DCount("*", "PriorEvaluationTrait", "PriorID=" & prior) = 7 And PriorEvaluationITA(prior) = "4.00", "Prior traits lost"
    Check RejectPreview(c, b), "Nested merge accepted"
    stamp = DirectoryMergeStamp(a, b)
    db.Execute "UPDATE PriorEvaluation SET VersionCounter=VersionCounter+1 WHERE PriorID=" & prior, dbFailOnError
    Check RejectUndo(merged), "Later prior edit did not block undo"
    db.Execute "UPDATE PriorEvaluation SET VersionCounter=VersionCounter-1 WHERE PriorID=" & prior, dbFailOnError
    db.Execute "UPDATE MemberReport SET MatchSource='Later decision' WHERE ReportID=" & report, dbFailOnError
    Check RejectUndo(merged), "Later link edit did not block undo"
    db.Execute "UPDATE MemberReport SET MatchSource='Reviewed directory merge' WHERE ReportID=" & report, dbFailOnError
    db.Execute "UPDATE Report SET CurrentRevisionID=" & CurrentRevision(otherReport) & " WHERE ReportID=" & report, dbFailOnError
    Check RejectUndo(merged), "Later report revision did not block undo"
    db.Execute "UPDATE Report SET CurrentRevisionID=" & rev & " WHERE ReportID=" & report, dbFailOnError
    Check DirectoryMergeStamp(a, b) = stamp, "Fixture restore mismatch"
    Check UndoDirectoryMerge(merged), "Undo failed"
    Check DLookup("MemberID", "MemberReport", "ReportID=" & report) = a And DLookup("MemberID", "PriorEvaluation", "PriorID=" & prior) = a, "Undo lost original IDs"
    Check Not DLookup("IsArchived", "CommandMember", "MemberID=" & a), "Original profile not recovered"
    Set originalLink = db.OpenRecordset("SELECT * FROM MemberReport WHERE ReportID=" & report, dbOpenSnapshot)
    For Each field In savedLink.Keys: Check SameFieldValue(originalLink.Fields(field), savedLink(field)), "Original link metadata not restored: " & field: Next
    originalLink.Close
    Set originalPrior = db.OpenRecordset("SELECT * FROM PriorEvaluation WHERE PriorID=" & prior, dbOpenSnapshot)
    For Each field In savedPrior.Keys: Check SameFieldValue(originalPrior.Fields(field), savedPrior(field)), "Original prior metadata not restored: " & field: Next
    originalPrior.Close
    Check DLookup("VersionCounter", "PriorEvaluation", "PriorID=" & prior) = 3, "Prior version not monotonic"
    Check DLookup("MergeState", "DirectoryMerge", "MergeID=" & merged) = "Undone" And RejectUndo(merged), "Undo history/idempotency failed"
    'Retain both committed and undone records for schema-27 transfer verification.
    merged = ApplyDirectoryMerge(a, b, PreviewDirectoryMerge(a, b), "Reviewed same Sailor after guarded undo")
    Check DCount("*", "DirectoryMerge") = 2, "Retained history lost"
    Dim x As Long, y As Long
    x = SaveDirectoryMember(0, 0, "Alex Noidentifier", "NOIDENTIFIER, ALEX", "IT2", "")
    y = SaveDirectoryMember(0, 0, "Alex Newname", "NEWNAME, ALEX", "IT2", "")
    Check Len(PreviewDirectoryMerge(x, y)) = 64, "Missing IDs/changed name prevented deliberate review"
    stamp = PreviewDirectoryMerge(x, y)
    DoCmd.OpenForm "frmUXMemberEditor"
    Check RejectApply(x, y, stamp), "Open member editor did not block merge"
    DoCmd.Close acForm, "frmUXMemberEditor", acSaveNo
    db.Execute "UPDATE CommandMember SET IsArchived=Null WHERE MemberID=" & x, dbFailOnError
    beforeArchive = DLookup("IsArchived", "CommandMember", "MemberID=" & x)
    merged = ApplyDirectoryMerge(x, y, PreviewDirectoryMerge(x, y), "Reviewed member provided records")
    Check UndoDirectoryMerge(merged), "Empty link merge undo failed"
    Check SameFieldValue(DLookup("IsArchived", "CommandMember", "MemberID=" & x), beforeArchive), "Original archive flag not restored exactly"
    db.Execute "UPDATE CommandMember SET IsArchived=False WHERE MemberID=" & x, dbFailOnError
    DoCmd.OpenForm "frmUXDirectoryMerge"
    Forms!frmUXDirectoryMerge!cboSource = x: Forms!frmUXDirectoryMerge!cboTarget = y
    Check UXDirectoryMergePreview(), "Embedded preview route failed"
    Dim owner As Long, second As Long
    owner = ActiveContributorID(): second = AddContributor("Second synthetic merger")
    Forms!frmUXDirectoryMerge!txtReason = "Synthetic same member review"
    SwitchContributor second
    Check Not UXDirectoryMergeApply(), "Another contributor applied the reviewed directory merge"
    Check Not Nz(DLookup("IsArchived", "CommandMember", "MemberID=" & x), False), "Rejected merge ownership changed member availability"
    SwitchContributor owner
    Check Not UXDirectoryMergeApply(), "Rejected merge ownership retained actionable preview"
    Check UXDirectoryMergePreview(), "Original contributor could not preview again"
    DoCmd.Close acForm, "frmUXDirectoryMerge", acSaveNo
    DoCmd.OpenForm "frmUXDirectoryMerge"
    Forms!frmUXDirectoryMerge!cboSource = x: Forms!frmUXDirectoryMerge!cboTarget = y
    Forms!frmUXDirectoryMerge!txtReason = "Synthetic reviewed duplicate"
    Check Not UXDirectoryMergeApply() And Not Forms!frmUXDirectoryMerge!cmdApply.Enabled, "Reopened directory merge reused a previous preview"
    Check UXDirectoryMergePreview(), "Reopened directory merge could not preview again"
    Check InStr(Forms!frmUXDirectoryMerge!txtPreview, "000000005") = 0, "Unmasked identifier in preview"
    Check UXWindowHelp("frmUXDirectoryMerge"), "Context help failed"
    Check InStr(Forms!frmUXHelp!txtGuide, "Duplicate Sailor review") > 0 And CurrentProject.AllForms("frmUXDirectoryMerge").IsLoaded, "Help did not retain the review window"
    DoCmd.Close acForm, "frmUXHelp", acSaveNo
    DoCmd.Close acForm, "frmUXDirectoryMerge", acSaveNo
    'Reporting senior links follow the surviving member and return on undo.
    Dim p As Long, q As Long, linked As Long, values As Object
    p = SaveDirectoryMember(0, 0, "Jordan Lee", "LEE, JORDAN", "LCDR", "")
    q = SaveDirectoryMember(0, 0, "Jordan Lee", "LEE, JORDAN A", "LCDR", "")
    Set values = RSSeniorValues(0): values("ReportingSeniorName") = "LEE, J"
    linked = RSSaveSenior(0, "Synthetic linked senior", values, 0, p)
    Check DLookup("MemberID", "ReportingSeniorRecord", "SeniorID=" & linked) = p, "Senior link not saved"
    DoCmd.OpenForm "frmUXSeniorEditor", , , , , , CStr(linked)
    Check RejectApply(p, q, PreviewDirectoryMerge(p, q)), "Open reporting senior editor did not block merge"
    DoCmd.Close acForm, "frmUXSeniorEditor", acSaveNo
    merged = ApplyDirectoryMerge(p, q, PreviewDirectoryMerge(p, q), "Reviewed linked senior duplicate")
    Check DLookup("MemberID", "ReportingSeniorRecord", "SeniorID=" & linked) = q And DCount("*", "DirectoryMergeSenior", "MergeID=" & merged) = 1, "Merge did not move the senior link"
    Check UndoDirectoryMerge(merged), "Undo of a linked senior merge failed"
    Check DLookup("MemberID", "ReportingSeniorRecord", "SeniorID=" & linked) = p, "Undo did not restore the senior link"
    Call RSSaveSenior(0, "Second linked senior", values, 0, q)
    Check RejectPreview(p, q), "Merge of two members that both back seniors accepted"
    TestDirectoryMerge = "PASS: " & checks & " directory merge checks; committed and undone fixtures retained": Exit Function
Failed:
    TestDirectoryMerge = "FAIL: " & Err.Number & ": " & Err.Description
End Function
