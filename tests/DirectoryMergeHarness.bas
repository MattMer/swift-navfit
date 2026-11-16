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
Private Function RosterFixture(ByVal db As DAO.Database, ByVal guid As String) As Long
    db.Execute "INSERT INTO RosterImport (ImportGuid,SourceFormat,AsOfDate,ImportedBy,ImportedAt) VALUES ('" & guid & "','MNA',#2026-06-01#," & MajorAuthor() & ",Now())", dbFailOnError
    RosterFixture = db.OpenRecordset("SELECT @@IDENTITY")(0)
End Function
Private Function DepartureFixture(ByVal db As DAO.Database, ByVal member As Long, ByVal state As String) As Long
    db.Execute "INSERT INTO MemberDeparture (MemberID,Basis,DepartureState,CreatedBy,CreatedAt,VersionCounter) VALUES (" & member & ",'PRD','" & state & "'," & MajorAuthor() & ",Now(),1)", dbFailOnError
    DepartureFixture = db.OpenRecordset("SELECT @@IDENTITY")(0)
End Function
Private Function DismissalFixture(ByVal db As DAO.Database, ByVal member As Long, ByVal due As Date) As Long
    db.Execute "INSERT INTO ReportDueDismissal (MemberID,Basis,DueDate,Reason,DismissedBy,DismissedAt) VALUES (" & member & ",'Periodic',#" & Format$(due, "yyyy-mm-dd") & "#,'Synthetic'," & MajorAuthor() & ",Now())", dbFailOnError
    DismissalFixture = db.OpenRecordset("SELECT @@IDENTITY")(0)
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
    'Roster presence, departures and due-date dismissals follow the survivor; an entry the
    'survivor already has stays with the merged member. Undo moves them back.
    Dim m As Long, n As Long, imp1 As Long, imp2 As Long, dep1 As Long, dep2 As Long, dis1 As Long, dis2 As Long
    Set db = CurrentDb
    m = SaveDirectoryMember(0, 0, "Riley Ames", "AMES, RILEY", "PO2", "")
    n = SaveDirectoryMember(0, 0, "Riley Ames", "AMES, RILEY B", "PO2", "")
    imp1 = RosterFixture(db, "{00000000-0000-0000-0000-0000000000A1}")
    imp2 = RosterFixture(db, "{00000000-0000-0000-0000-0000000000A2}")
    db.Execute "INSERT INTO RosterPresence (RosterImportID,MemberID,UIC) VALUES (" & imp1 & "," & m & ",'00001')", dbFailOnError
    db.Execute "INSERT INTO RosterPresence (RosterImportID,MemberID,UIC) VALUES (" & imp2 & "," & m & ",'00001')", dbFailOnError
    db.Execute "INSERT INTO RosterPresence (RosterImportID,MemberID,UIC) VALUES (" & imp2 & "," & n & ",'00001')", dbFailOnError
    dep1 = DepartureFixture(db, m, "Pending")
    dep2 = DepartureFixture(db, n, "Pending")
    dis1 = DismissalFixture(db, m, #6/30/2026#)
    dis2 = DismissalFixture(db, m, #9/30/2026#)
    Call DismissalFixture(db, n, #9/30/2026#)
    merged = ApplyDirectoryMerge(m, n, PreviewDirectoryMerge(m, n), "Reviewed roster duplicate")
    Check DCount("*", "RosterPresence", "MemberID=" & n) = 2 And DCount("*", "RosterPresence", "MemberID=" & m & " AND RosterImportID=" & imp2) = 1, "Roster presence did not follow the survivor"
    Check DLookup("MemberID", "MemberDeparture", "DepartureID=" & dep1) = m, "An open departure was moved onto a survivor that has one"
    Check DLookup("MemberID", "ReportDueDismissal", "DismissalID=" & dis1) = n And DLookup("MemberID", "ReportDueDismissal", "DismissalID=" & dis2) = m, "Dismissals did not follow the survivor"
    Check DCount("*", "DirectoryMergeLink", "MergeID=" & merged) = 2, "Moved roster entries not recorded"
    Check UndoDirectoryMerge(merged), "Undo of a roster merge failed"
    Check DCount("*", "RosterPresence", "MemberID=" & m) = 2 And DCount("*", "RosterPresence", "MemberID=" & n) = 1 And DLookup("MemberID", "ReportDueDismissal", "DismissalID=" & dis1) = m, "Undo did not restore roster entries"
    db.Execute "UPDATE MemberDeparture SET DepartureState='Returned' WHERE DepartureID=" & dep2, dbFailOnError
    merged = ApplyDirectoryMerge(m, n, PreviewDirectoryMerge(m, n), "Reviewed roster duplicate again")
    Check DLookup("MemberID", "MemberDeparture", "DepartureID=" & dep1) = n, "Open departure did not follow a survivor without one"
    Check UndoDirectoryMerge(merged) And DLookup("MemberID", "MemberDeparture", "DepartureID=" & dep1) = m, "Undo did not restore the departure"
    TestDirectoryMerge = "PASS: " & checks & " directory merge checks; committed and undone fixtures retained": Exit Function
Failed:
    TestDirectoryMerge = "FAIL: " & Err.Number & ": " & Err.Description
End Function
'Directory-wide possible duplicates (schema 51): strength, conflicts, survivor choice,
'saved "different Sailors" decisions and the review window.
Public Function TestDuplicatePairs() As String
    Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long, p5 As Long, p6 As Long, p7 As Long, item As Object, dismissal As Long, f As Access.Form, row As Long
    On Error GoTo Failed
    p1 = SaveDirectoryMember(0, 0, "Jordan Quill", "QUILL, JORDAN", "IT1", "")
    p2 = SaveDirectoryMember(0, 0, "Jordan Quill", "QUILL, JORDAN", "IT1", "")
    p3 = SaveDirectoryMember(0, 0, "Jordan Quill", "QUILL, JORDAN", "IT1", "000000101")
    p4 = SaveDirectoryMember(0, 0, "Jordan Quill", "QUILL, JORDAN", "IT1", "000000102")
    p5 = SaveDirectoryMember(0, 0, "Jo Quill", "QUILL, JO", "IT2", "")
    'Two profiles never share an SSN, but their reports can carry the same DoD ID.
    p6 = SaveDirectoryMember(0, 0, "Riley Vance", "VANCE, RILEY", "YN1", "")
    p7 = SaveDirectoryMember(0, 0, "Morgan Ashby", "ASHBY, MORGAN", "YN1", "")
    LinkedReportDoDID p6, "Riley Vance", "1234500201"
    LinkedReportDoDID p7, "Morgan Ashby", "1234500201"
    Set item = FindPair(DirectoryDuplicatePairs(), p6, p7)
    Check Not item Is Nothing, "Same DoD ID not listed"
    Check item("Strength") = 1 And InStr(item("Why"), "Same DoD ID") > 0 And InStr(item("Why"), "201") = 0, "Same DoD ID not the strongest match, or identifier shown"
    Set item = FindPair(DirectoryDuplicatePairs(), p1, p2)
    Check Not item Is Nothing, "Same report name not listed"
    Check item("Strength") = 2 And InStr(item("Why"), "QUILL, JORDAN") > 0, "Same report name strength"
    Check FindPair(DirectoryDuplicatePairs(), p3, p4) Is Nothing, "Members with different SSNs listed as duplicates"
    Check Not FindPair(DirectoryDuplicatePairs(), p1, p3) Is Nothing, "Profile without an SSN not paired with a namesake"
    Set item = FindPair(DirectoryDuplicatePairs(), p1, p5)
    Check Not item Is Nothing, "Same surname and initial not listed"
    Check item("Strength") = 3, "Surname and initial strength"
    Check FindPair(DirectoryDuplicatePairs(), p6, p1) Is Nothing, "Unrelated members listed"
    Check DuplicateSurvivor(p1, p3) = p3 And DuplicateSurvivor(p3, p1) = p3, "Survivor is not the SSN holder"
    Check DuplicateSurvivor(p1, p2) = p1, "Survivor tie is not the older profile"
    Check RejectDismiss(p1, p2, ""), "Dismissal without a reason accepted"
    Check RejectDismiss(p1, p2, "SSN 000-00-0101 differs"), "Dismissal reason with an identifier accepted"
    Check RejectDismiss(p1, p1, "Same member"), "Dismissal of one member accepted"
    dismissal = DismissDuplicatePair(p2, p1, "Different Sailors, checked with the LPO")
    Check dismissal > 0 And DLookup("LowMemberID", "MemberDuplicateDismissal", "DismissalID=" & dismissal) = p1, "Dismissal not stored low/high"
    Check FindPair(DirectoryDuplicatePairs(), p1, p2) Is Nothing, "Dismissed pair still listed"
    Set item = FindPair(DirectoryDuplicatePairs(True), p1, p2)
    Check Not item Is Nothing, "Dismissed pair not shown on request"
    Check item("DismissalID") = dismissal, "Dismissed pair lost its dismissal"
    Check Not FindPair(DirectoryDuplicatePairs(), p1, p3) Is Nothing, "Dismissing one pair hid another"
    Check RejectDismiss(p1, p2, "Again"), "Second dismissal of the same pair accepted"
    RestoreDuplicatePair dismissal
    Check Not FindPair(DirectoryDuplicatePairs(), p1, p2) Is Nothing, "Restored pair not listed"
    Check Not IsNull(DLookup("UndoneAt", "MemberDuplicateDismissal", "DismissalID=" & dismissal)), "Restore erased the dismissal history"
    Check RejectRestore(dismissal), "Second restore accepted"
    'The review window.
    DoCmd.OpenForm "frmUXDirectoryMerge": Set f = Forms!frmUXDirectoryMerge
    Check InStr(f!lblPairs.Caption, "Possible duplicate Sailors:") > 0, "Pair count not shown"
    row = PairRow(f, p6, p7)
    Check row > 0, "Same SSN pair not in the window"
    Check f!lstPairs.Column(1, row) = "Likely" And PairRow(f, p1, p5) > row, "Window not ordered strongest first"
    f!lstPairs = f!lstPairs.Column(0, row)
    Check UXDirectoryMergePair(), "Selecting a pair failed"
    Check f!cboTarget = p6 And f!cboSource = p7 And f!cmdDismiss.Enabled And Not f!cmdRestore.Enabled, "Selecting a pair did not fill both members"
    f!txtReason = "Synthetic different Sailors"
    Check UXDirectoryMergeDismiss(), "Window dismissal failed"
    Check PairRow(f, p6, p7) = 0 And IsNull(f!cboSource) And IsNull(f!txtReason), "Window still lists the dismissed pair"
    f!chkShowDismissed = True: Call UXDirectoryMergePairs
    row = PairRow(f, p6, p7)
    Check row > 0, "Window does not show dismissed pairs on request"
    Check f!lstPairs.Column(1, row) = "Different", "Dismissed pair not labeled"
    f!lstPairs = f!lstPairs.Column(0, row): Call UXDirectoryMergePair
    Check f!cmdRestore.Enabled And Not f!cmdDismiss.Enabled, "Dismissed pair offers the wrong action"
    Check UXDirectoryMergeRestore(), "Window restore failed"
    f!chkShowDismissed = False: Call UXDirectoryMergePairs
    Check PairRow(f, p6, p7) > 0, "Restored pair not back in the window"
    DoCmd.Close acForm, "frmUXDirectoryMerge", acSaveNo
    TestDuplicatePairs = "PASS: " & checks & " possible-duplicate list, dismissal and window checks": Exit Function
Failed:
    TestDuplicatePairs = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
Private Sub LinkedReportDoDID(ByVal member As Long, ByVal name As String, ByVal dodid As String)
    Dim report As Long, changes As Object
    DirectoryCreationContext member
    report = CommitDraft(0, "EVAL", "Duplicate fixture", name, #1/1/2026#, #9/30/2026#, "", "Test", "")
    DirectoryCreationContext 0
    'Saved as a revision, so the fixture workspace stays valid for the release import check.
    Set changes = CreateObject("Scripting.Dictionary"): changes.Add "MemberDoDID", dodid
    Call CommitSavedFields(report, CurrentRevision(report), changes, "Duplicate fixture DoD ID")
    Check DLookup("MemberDoDID", "Report", "ReportID=" & report) = dodid, "Fixture DoD ID not saved"
    Check DLookup("MemberID", "MemberReport", "ReportID=" & report) = member, "Fixture report not linked to its member"
End Sub
Private Function FindPair(ByVal pairs As Collection, ByVal a As Long, ByVal b As Long) As Object
    Dim item As Object
    For Each item In pairs
        If (item("LowID") = a And item("HighID") = b) Or (item("LowID") = b And item("HighID") = a) Then Set FindPair = item: Exit Function
    Next
End Function
Private Function PairRow(ByVal f As Access.Form, ByVal a As Long, ByVal b As Long) As Long
    Dim i As Long, low As Long, high As Long
    If a < b Then low = a: high = b Else low = b: high = a
    For i = 1 To f!lstPairs.ListCount - 1
        If Split(f!lstPairs.Column(0, i) & "--", "-")(0) = CStr(low) And Split(f!lstPairs.Column(0, i) & "--", "-")(1) = CStr(high) Then PairRow = i: Exit Function
    Next
End Function
Private Function RejectDismiss(ByVal a As Long, ByVal b As Long, ByVal reason As String) As Boolean
    On Error GoTo Rejected
    Call DismissDuplicatePair(a, b, reason): Exit Function
Rejected:
    RejectDismiss = True
End Function
Private Function RejectRestore(ByVal id As Long) As Boolean
    On Error GoTo Rejected
    RestoreDuplicatePair id: Exit Function
Rejected:
    RejectRestore = True
End Function
