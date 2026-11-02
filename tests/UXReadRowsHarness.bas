Attribute VB_Name = "UXReadRowsHarness"
Option Compare Binary
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 8750, , message
    checks = checks + 1
End Sub
Private Sub CompareReports()
    Dim r As DAO.Recordset, before As Long, snapshots As Long, id As Long
    before = DCount("*", "ReportRevision"): snapshots = DCount("*", "GroupSnapshot") + DCount("*", "ComparisonSet")
    UXPrepareReportRows
    Set r = CurrentDb.OpenRecordset("SELECT ReportID,CurrentRevisionID FROM Report", dbOpenSnapshot)
    Do While Not r.EOF
        id = r!ReportID
        Check SameFieldValue(UXReportRowCell(id, 0), UXMemberLabel(id)), "Member differs"
        Check SameFieldValue(UXReportRowCell(id, 1), UXReportState(id)), "State differs for " & id
        Check SameFieldValue(UXReportRowCell(id, 2), UXTableITA(id)), "ITA differs for " & id
        Check SameFieldValue(UXReportRowCell(id, 3), UXTableGroups(id)), "Groups differ"
        Check SameFieldValue(UXReportRowCell(id, 4), DLookup("RevisionNumber", "ReportRevision", "RevisionID=" & r!CurrentRevisionID)), "Revision number differs"
        r.MoveNext
    Loop
    r.Close
    Check DCount("*", "ReportRevision") = before, "Reading created revisions"
    Check DCount("*", "GroupSnapshot") + DCount("*", "ComparisonSet") = snapshots, "Reading created historical snapshots"
End Sub
' Quoted value-list cells keep commas, semicolons and quotes in their own column,
' as the NAVFIT group picker, report picker and NavyWriter list rely on.
Private Sub CheckListCells()
    Dim l As Access.ListBox, title As String
    title = "Alpha, Bravo; ""Charlie"" Group"
    DoCmd.OpenForm "frmUXNativeGroups", , , , , acHidden: Set l = Forms!frmUXNativeGroups!lstGroups
    l.RowSourceType = "Value List": l.RowSource = ""
    l.AddItem "7;" & UXListCell(title) & ";" & UXListCell("EVAL") & ";" & UXListCell("E6") & ";3"
    l.AddItem "8;" & UXListCell("SMITH, JOHN Q") & ";" & UXListCell("FITREP") & ";" & UXListCell("O3") & ";12"
    Check l.ListCount = 2, "Quoted cells changed the row count"
    Check l.Column(1, 0) = title And l.Column(2, 0) = "EVAL" And l.Column(3, 0) = "E6" And l.Column(4, 0) = "3", "Punctuation shifted list columns"
    Check l.Column(0, 1) = "8" And l.Column(1, 1) = "SMITH, JOHN Q" And l.Column(4, 1) = "12", "Comma name shifted list columns"
    DoCmd.Close acForm, "frmUXNativeGroups", acSaveNo
End Sub
Private Function RejectDivergedMember(ByVal groupID As Long, ByVal reportID As Long) As Boolean
    On Error GoTo Rejected
    AddGroupMember groupID, reportID, "E5"
    Exit Function
Rejected:
    RejectDivergedMember = InStr(Err.Description, "Derived report cache differs") > 0
End Function
Public Function TestUXReadRows() As String
    Dim id As Long, other As Long, p As Long, g As Long, g2 As Long, template As Long, run As Long, task As Long, thread As Long
    Dim changes As Object, key As Variant, revision As Long, f As Access.Form, preset As Variant, col As Long, rows As Long, signed As Long
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC READ ROWS", "Reader"
    CheckListCells
    p = MajorAuthor()
    id = CommitDraft(0, "EVAL", "Z title", "ZULU, MEMBER", #1/1/2026#, #9/30/2026#, "")
    other = CommitDraft(0, "EVAL", "A title", "ALPHA, MEMBER", Null, Null, "")
    CompareReports
    Set changes = CreateObject("Scripting.Dictionary")
    For Each key In TraitKeys(): changes(CStr(key)) = "4.0": Next
    changes("OfficialName") = "SAVED, NAME"
    CommitSavedFields id, CurrentRevision(id), changes, "Complete traits"
    CompareReports
    Check UXReportRowCell(id, 2) = "4.00", "Complete traits average"
    ' Scratch-only divergence proves calculations use saved values, not live Report traits.
    CurrentDb.Execute "UPDATE Report SET Trait33='1.0' WHERE ReportID=" & id, dbFailOnError
    CompareReports
    Check UXReportRowCell(id, 2) = "4.00", "Live traits replaced saved revision"
    revision = CurrentRevision(id)
    CurrentDb.Execute "UPDATE RevisionValue SET IsNullValue=True WHERE RevisionID=" & revision & " AND FieldKey='Trait33'", dbFailOnError
    CurrentDb.Execute "UPDATE RevisionValue SET ValueText='NOB' WHERE RevisionID=" & revision & " AND FieldKey='Trait34'", dbFailOnError
    CurrentDb.Execute "UPDATE RevisionValue SET ValueText='' WHERE RevisionID=" & revision & " AND FieldKey='Trait35'", dbFailOnError
    CompareReports
    Check UXReportRowCell(id, 2) = "4.00", "Partial trait semantics changed"
    CurrentDb.Execute "DELETE FROM RevisionValue WHERE RevisionID=" & revision & " AND FieldKey='Trait39'", dbFailOnError
    CompareReports
    Check UXReportRowCell(id, 2) = "Unavailable", "Missing historical field accepted"
    CurrentDb.Execute "UPDATE Report SET ObservationStatus='NOB' WHERE ReportID=" & id, dbFailOnError
    CompareReports
    Check UXReportRowCell(id, 2) = "NOB", "NOB override lost"
    CurrentDb.Execute "INSERT INTO RevisionValue (RevisionID,FieldKey,ValueText,IsNullValue) VALUES (" & revision & ",'Trait39','4.0',False)", dbFailOnError
    g = CreateSummaryGroup("Z group", "EVAL", "E5"): g2 = CreateSummaryGroup("A; quoted group", "EVAL", "E5")
    Check RejectDivergedMember(g, id), "Membership accepted the deliberately diverged report cache"
    Check DCount("*", "GroupMember", "GroupID=" & g) = 0, "Rejected divergence changed membership"
    'The preceding corrupt-cache cases are read-only probes. Restore their exact
    'seed values before testing unrelated guarded review/group mutations.
    CurrentDb.Execute "UPDATE Report SET Trait33='4.0',ObservationStatus=Null WHERE ReportID=" & id, dbFailOnError
    CurrentDb.Execute "UPDATE RevisionValue SET ValueText='4.0',IsNullValue=False WHERE RevisionID=" & revision & " AND FieldKey IN ('Trait33','Trait34','Trait35')", dbFailOnError
    AddGroupMember g, id, "E5": AddGroupMember g2, id, "E5"
    CompareReports
    Check UXReportRowCell(id, 3) = "A; quoted group; Z group", "Group ordering/semicolons changed"
    task = AssignReview(id, CurrentRevision(id), p, Null, "Review")
    CompareReports
    Check UXReportRowCell(id, 1) = "In review", "Pending assignment state"
    FinishReview task, "Returned", "Revise"
    CompareReports
    Check UXReportRowCell(id, 1) = "Returned", "Returned assignment state"
    template = CreateRouteTemplate("Read route", CStr(p), "Review saved values")
    run = StartRoute(id, template, g)
    CompareReports
    Check UXReportRowCell(id, 1) = "In review", "Route must outrank returned assignment"
    AddGroupMember g, other, "E5"
    CompareReports
    Check UXReportRowCell(id, 1) = "Review needs restart", "Group membership staleness lost"
    run = RestartStaleRoute(run, CurrentRevision(id), GroupStamp(g))
    DecideRouteStep run, "Accepted", "Accepted"
    CompareReports
    ApproveSavedRoute run, "Approved"
    CompareReports
    Check UXReportRowCell(id, 1) = "Local approval", "Approval precedence"
    ArchiveDraft id, True
    CompareReports
    Check UXReportRowCell(id, 1) = "Archived", "Archive precedence"
    ArchiveDraft id, False
    DoCmd.OpenForm "frmLibrary": Set f = Forms!frmLibrary
    For Each preset In Array("Drafting", "Dates", "History")
        SaveLibraryView CStr(preset): RefreshLibrary
        Check f!lstReports.ListCount = 3, "Preset omitted report"
        Check f!lstReports.ColumnCount = 10 And f!lstReports.Column(4, 0) = "Rate", "Preset column shape"
        For rows = 1 To 2
            Check SameFieldValue(f!lstReports.Column(7, rows), UXTableITA(CLng(f!lstReports.Column(0, rows)))), "Displayed ITA mismatch"
        Next
    Next
    f!uxStatusFilter = "Local approval": RefreshLibrary
    Check f!lstReports.ListCount = 2 And f!lstReports.Column(0, 1) = id, "Cached status filter differs"
    f!uxStatusFilter = "All": f!lstReports = other: RefreshLibrary
    Check f!lstReports = other, "Refresh selection lost"
    Check UXSortList("frmLibrary", "lstReports", 6), "Prepared ITA sort failed"
    Check f!lstReports = other, "Sort selection lost"
    DoCmd.Close acForm, "frmLibrary", acSaveNo
    ' Signed and reopened reports outrank review overlays in both readers, and
    ' review work on a signed report leaves My Work until it is reopened.
    signed = CommitDraft(0, "EVAL", "Signed title", "SIERRA, MEMBER", #1/1/2026#, #9/30/2026#, "")
    task = AssignReview(signed, CurrentRevision(signed), p, Null, "Review before signing")
    run = StartRoute(signed, template)
    MarkReportSigned signed, CurrentRevision(signed), Date
    CompareReports
    Check UXReportRowCell(signed, 1) = "Signed", "Review overlay replaced Signed"
    Check Not UXWorkVisible("Assignment", task, p) And Not UXWorkVisible("Route", run, p), "Signed report review stayed in My Work"
    Check Not UXWorkVisibleRows(p).Exists("Assignment:" & task) And Not UXWorkVisibleRows(p).Exists("Route:" & run), "Signed report review stayed in batched My Work"
    ReopenSignedReport signed, "Correct block 43"
    CompareReports
    Check UXReportRowCell(signed, 1) = "Reopened for correction", "Reopened report shown as " & UXReportRowCell(signed, 1)
    Check UXWorkVisible("Assignment", task, p) And UXWorkVisibleRows(p).Exists("Assignment:" & task), "Reopened report review missing from My Work"
    MarkReportSigned signed, CurrentRevision(signed), Date
    CompareReports
    Check UXReportRowCell(signed, 1) = "Signed", "Re-signed report state"
    TestUXReadRows = "PASS: " & checks & " batch report equivalence, historical traits, review state, filters and sorting checks": Exit Function
Failed:
    TestUXReadRows = "FAIL after " & checks & ": " & Err.Number & " " & Err.Description
End Function
