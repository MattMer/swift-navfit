Attribute VB_Name = "O3RegressionHarness"
Option Explicit
Public Function TestO3Cells() As String
 Dim rows As Access.Form, r As DAO.Recordset, id As Variant, result As String, mark As String
 On Error GoTo Failed
 DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!cboGroup = 3
 UXWorkbenchMode "Members": UXOpenComparison
 Set rows = Forms!frmUXComparison!uxTraitGrid.Form
 For Each id In Array(6,7)
  Set r = rows.RecordsetClone: r.FindFirst "ReportID=" & id: rows.Bookmark = r.Bookmark: r.Close
  mark = IIf(Nz(rows!Trait33, "") = "2.0", "3.0", "2.0"): rows!Trait33 = mark
  If Not UXGridEdit("Trait33", vbYes) Then Err.Raise vbObjectError + 7990, , Forms!frmUXComparison!lblStatus.Caption
  If DLookup("Trait33", "Report", "ReportID=" & id) <> mark Then Err.Raise vbObjectError + 7990, , "Saved trait reverted"
  rows!Trait34 = "5.0"
  If Not UXGridEdit("Trait34", vbYes) Then Err.Raise vbObjectError + 7990, , Forms!frmUXComparison!lblStatus.Caption
  If DLookup("Trait34", "Report", "ReportID=" & id) <> "5.0" Then Err.Raise vbObjectError + 7990, , "Second save reverted"
  result = result & id & " repeated saves persisted; "
 Next
 UXCloseComparison: UXOpenComparison
 Set rows = Forms!frmUXComparison!uxTraitGrid.Form
 For Each id In Array(6,7)
  Set r = rows.RecordsetClone: r.FindFirst "ReportID=" & id: rows.Bookmark = r.Bookmark: r.Close
  If rows!Trait34 <> "5.0" Then Err.Raise vbObjectError + 7990, , "Reopen lost trait"
 Next
 TestO3Cells = "PASS: " & result & "reopen verified": Exit Function
Failed:
 TestO3Cells = "FAIL: " & Err.Number & " " & Err.Description
End Function
