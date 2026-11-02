Attribute VB_Name = "O3RegressionHarness"
Option Explicit
Public Function TestO3Cells() As String
 Dim rows As Access.Form, r As DAO.Recordset, id As Variant, result As String, mark As String, g As Long, ids(1) As Long
 On Error GoTo Failed
 ' Synthetic O3 FITREP group with two members replaces the original incident workspace.
 ConfigureWorkspace "SYNTHETIC O3 CELLS", "Tester"
 g = CreateSummaryGroup("O3 cells", "FITREP", "O3")
 ids(0) = CommitDraft(0, "FITREP", "O3 cells", "Synthetic officer one", #1/1/2026#, #9/30/2026#, "", "Test", "")
 ids(1) = CommitDraft(0, "FITREP", "O3 cells", "Synthetic officer two", #1/1/2026#, #9/30/2026#, "", "Test", "")
 AddGroupMember g, ids(0), "O3": AddGroupMember g, ids(1), "O3"
 DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = g
 UXWorkbenchMode "Members": UXOpenComparison
 Set rows = Forms!frmUXComparison!uxTraitGrid.Form
 For Each id In Array(ids(0), ids(1))
  Set r = rows.RecordsetClone: r.FindFirst "ReportID=" & id: rows.Bookmark = r.Bookmark: r.Close
  mark = IIf(Nz(rows!Trait33, "") = "2.0", "3.0", "2.0"): rows!Trait33 = mark
  If Not UXGridEdit("Trait33", vbYes) Then Err.Raise vbObjectError + 7990, , Forms!frmUXComparison!lblStatus.Caption
  If DLookup("Trait33", "Report", "ReportID=" & id) <> mark Then Err.Raise vbObjectError + 7990, , "Saved trait reverted"
  rows!Trait34 = "5.0"
  If Not UXGridEdit("Trait34", vbYes) Then Err.Raise vbObjectError + 7990, , Forms!frmUXComparison!lblStatus.Caption
  If DLookup("Trait34", "Report", "ReportID=" & id) <> "5.0" Then Err.Raise vbObjectError + 7990, , "Second save reverted"
  result = result & id & " repeated saves persisted; "
 Next
 Call UXCloseComparison: Call UXShowGroupMembers(g)
 Set rows = Forms!frmUXComparison!uxTraitGrid.Form
 For Each id In Array(ids(0), ids(1))
  Set r = rows.RecordsetClone: r.FindFirst "ReportID=" & id: rows.Bookmark = r.Bookmark: r.Close
  If rows!Trait34 <> "5.0" Then Err.Raise vbObjectError + 7990, , "Reopen lost trait"
 Next
 TestO3Cells = "PASS: " & result & "reopen verified": Exit Function
Failed:
 TestO3Cells = "FAIL: " & Err.Number & " " & Err.Description
End Function
