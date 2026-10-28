Attribute VB_Name = "PendingCellHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 7970, , message
 checks = checks + 1
End Sub
Private Sub SelectRow(ByVal id As Long)
 Dim f As Access.Form, r As DAO.Recordset
 Set f = Forms!frmUXComparison!uxTraitGrid.Form
 If f.Dirty Then f.Dirty = False
 Set r = f.RecordsetClone: r.FindFirst "ReportID=" & id: f.Bookmark = r.Bookmark: r.Close
End Sub
Private Function TrySwitch(ByVal profile As Long) As Long
 Dim result As Boolean
 On Error GoTo Failed
 result = SwitchContributor(profile): Exit Function
Failed:
 TrySwitch = Err.Number
End Function
Public Function TestPendingCells() As String
 Dim groupID As Long, first As Long, second As Long, otherProfile As Long, author As Long, revision As Long, count As Long, secondRevision As Long, path As String, f As Access.Form, db As DAO.Database
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC PENDING CELL TEST", "Original contributor"
 author = MajorAuthor(): otherProfile = AddContributor("Other synthetic contributor")
 groupID = CreateSummaryGroup("Pending cells", "EVAL", "E6")
 first = CreateGroupDraft(groupID, GroupStamp(groupID), UXDefaultsStamp(groupID), 0, 0, "Synthetic first", "FIRST, SYNTHETIC", "IT1", "", Null, Null)
 second = CreateGroupDraft(groupID, GroupStamp(groupID), UXDefaultsStamp(groupID), 0, 0, "Synthetic second", "SECOND, SYNTHETIC", "IT1", "", Null, Null)
 Set db = CurrentDb: db.TableDefs("Report").Fields("Trait35").ValidationRule = "<> '5.0'"
 DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = groupID
 Check UXOpenComparison(), "Open comparison failed"
 Set f = Forms!frmUXComparison!uxTraitGrid.Form
 SelectRow first: revision = CurrentRevision(first): secondRevision = CurrentRevision(second)
 f!Trait33 = "1.0": Check UXGridQueue("Trait33"), "Queue first cell"
 f!AssignmentRecommendationA = String$(21, "X"): Check UXGridQueue("AssignmentRecommendationA"), "Queue invalid cell"
 SelectRow second: f!Trait34 = "2.0": Check UXGridQueue("Trait34"), "Queue later cell"
 Check Not UXGridDrain(), "Invalid cell accepted"
 Check UXGridPendingCount() = 2, "Uncommitted queue not retained"
 Check DLookup("Trait33", "Report", "ReportID=" & first) = "1.0", "Earlier successful cell lost"
 Check CurrentRevision(second) = secondRevision, "Later cell unexpectedly saved"
 SelectRow first: Check f!AssignmentRecommendationA = String$(21, "X"), "Failed value not displayed"
 SelectRow second: Check f!Trait34 = "2.0", "Later pending value not displayed"
 Check Forms!frmUXComparison!cmdRetryCells.Visible And Forms!frmUXComparison!cmdDiscardCells.Visible, "Recovery actions hidden"
 path = UXGridCacheFile()
 Check Not UXCloseComparison(), "Failed close reported success"
 Check CurrentProject.AllForms("frmUXComparison").IsLoaded, "Close dropped pending edits"
 Check Len(Dir$(path)) > 0 And UXGridPendingCount() = 2, "Close destroyed recovery cache"
 Check Not UXGroupNavigate("frmUXComparison", 1), "Group navigation bypassed failed close"
 Check CurrentProject.AllForms("frmUXComparison").IsLoaded And UXGridPendingCount() = 2, "Navigation destroyed pending edits"
 Check Not UXApplyComparison(), "Proposals applied with pending cells"
 Check TrySwitch(otherProfile) = vbObjectError + 4815, "Contributor switched with pending edits"
 Check MajorAuthor() = author, "Blocked switch changed contributor"
 SelectRow first: f!AssignmentRecommendationA = "VALID": Check UXGridQueue("AssignmentRecommendationA"), "Correction not queued"
 Forms!frmUXComparison!cmdRetryCells.SetFocus
 Check UXGridDrain(), "Corrected retry failed"
 Check UXGridPendingCount() = 0, "Retry left cells queued"
 Check DLookup("AssignmentRecommendationA", "Report", "ReportID=" & first) = "VALID" And DLookup("Trait34", "Report", "ReportID=" & second) = "2.0", "Retry lost later cell"
 Check DCount("*", "ReportRevision", "ReportID=" & first & " AND RevisionID>" & revision) = 2, "Retry duplicated successful revision"
 Check DLookup("AuthorProfileID", "ReportRevision", "RevisionID=" & CurrentRevision(second)) = author, "Retry changed attribution"
 count = DCount("*", "ReportRevision"): Check UXGridDrain() And DCount("*", "ReportRevision") = count, "Empty retry created revisions"
 Check Not Forms!frmUXComparison!cmdRetryCells.Visible, "Recovery actions not cleared"
 SelectRow first: revision = CurrentRevision(first): f!Trait35 = "5.0": Check UXGridQueue("Trait35"), "Queue injected save failure"
 SelectRow second: secondRevision = CurrentRevision(second): f!Trait36 = "2.0": Check UXGridQueue("Trait36"), "Queue after save failure"
 Check Not UXGridDrain(), "Injected database failure accepted"
 Check CurrentRevision(first) = revision And CurrentRevision(second) = secondRevision, "Failed transaction changed revisions"
 Check UXGridPendingCount() = 2, "Database failure dropped queue"
 Check Not UXCloseComparison() And CurrentProject.AllForms("frmUXComparison").IsLoaded, "Database failure did not cancel native unload"
 Forms!frmUXComparison!cmdDiscardCells.SetFocus
 Check UXGridDiscardPending(), "Explicit discard failed"
 Check UXGridPendingCount() = 0, "Discard left pending cells"
 Check CurrentRevision(first) = revision And CurrentRevision(second) = secondRevision, "Discard changed saved revisions"
 Set f = Forms!frmUXComparison!uxTraitGrid.Form
 SelectRow first: Check SameFieldValue(f!Trait35.Value, MajorRevisionValue(CurrentRevision(first), "Trait35")), "Discard retained rejected field: displayed=" & Nz(f!Trait35, "<null>") & "; saved=" & Nz(MajorRevisionValue(CurrentRevision(first), "Trait35"), "<null>")
 SelectRow second: Check SameFieldValue(f!Trait36.Value, MajorRevisionValue(CurrentRevision(second), "Trait36")), "Discard retained later unsaved field"
 'A read-only cache file cannot be deleted: report it and keep its path for retry.
 path = UXGridCacheFile(): SetAttr path, vbReadOnly
 Check UXCloseComparison(), "Close after discard failed"
 Check Not CurrentProject.AllForms("frmUXComparison").IsLoaded, "Close after discard left form open"
 Check Len(Dir$(path)) > 0 And UXGridCacheFile() = path, "Undeletable cache path was forgotten"
 Check SwiftLastFailureContext()("Procedure") = "UXGridClose" And SwiftLastFailureContext()("Stage") = "Delete trait cache" And SwiftLastFailureContext()("Number") <> 0, "Cache removal failure lost its diagnostic"
 SetAttr path, vbNormal
 Check UXGridClose() And UXGridCacheFile() = "", "Cache removal retry failed"
 Check Len(Dir$(path)) = 0, "Closed cache was not removed"
 Check SwitchContributor(otherProfile), "Contributor blocked after pending edits cleared"
 TestPendingCells = "PASS: " & checks & " pending-cell preservation, partial save, retry, database failure, native close cancellation, discard and attribution assertions": Exit Function
Failed:
 TestPendingCells = "FAIL: " & Err.Number & " " & Err.Description
End Function
