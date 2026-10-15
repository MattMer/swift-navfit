Attribute VB_Name = "ImportRankHarness"
Option Explicit
Private n As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
 If Not ok Then Err.Raise vbObjectError + 8380, , message
 n = n + 1
End Sub
Public Function TestImportRank(ByVal pdf As String) As String
 Dim values As Object, key As Variant, id As Long, groupID As Long, f As Access.Form, screen As Variant, names As Variant
 On Error GoTo Failed
 ConfigureWorkspace "SYNTHETIC RANK FILTER", "Tester"
 Check NAVFITDecode("26JAN01", "counseling") = "2026-01-01", "Default native year did not use 20xx"
 Check NAVFITDecode("99DEC31", "counseling") = "2099-12-31", "Native year used OS pivot"
 Set values = PreviewStructuredPDF(pdf): Set values = PDFMappedValues("", "SSN")
 Check values("PeriodStart") = #1/1/2026#, "PDF required a century prefix"
 PDFClearPreview
 For Each screen In Array("frmPDFImport", "frmUXNativeGroups", "frmImportReview", "frmNAVFITExchange")
  DoCmd.OpenForm CStr(screen), acDesign
  Set f = Forms(CStr(screen))
  For Each key In Array("cboCentury", "cboNAVFITCentury", "cboReturnCentury")
   Check Not UXHasControl(f, CStr(key)), "Century input retained on " & screen
  Next
  DoCmd.Close acForm, CStr(screen), acSaveNo
 Next
 Check UXDirectoryBatchOpen(), "Open empty batch picker"
 Check Forms!frmUXDirectoryBatch!lstMembers.ListCount = 0, "Empty directory filter failed"
 DoCmd.Close acForm, "frmUXDirectoryBatch", acSaveNo
 id = SaveDirectoryMember(0, 0, "Synthetic IT1", "SYNTHETIC, IT1", "IT1", "")
 SaveDirectoryMember 0, 0, "Synthetic YN1", "SYNTHETIC, YN1", "YN1", ""
 SaveDirectoryMember 0, 0, "Synthetic YN2", "SYNTHETIC, YN2", "YN2", ""
 SaveDirectoryMember 0, 0, "Synthetic LT", "SYNTHETIC, LT", "LT", ""
 SaveDirectoryMember 0, 0, "Synthetic no rank", "SYNTHETIC, NONE", "", ""
 Check UXDirectoryBatchOpen(), "Open populated batch picker"
 Set f = Forms!frmUXDirectoryBatch
 Check f!lstMembers.ListCount = 5, "All ranks filter lost members"
 f!cboRankFilter = "E6": Check UXDirectoryRankFilter(), "Paygrade filter failed"
 Check f!lstMembers.ListCount = 2, "E6 did not include both IT1 and YN1"
 f!cboRankFilter = "YN1": UXDirectoryRankFilter
 Check f!lstMembers.ListCount = 1, "Specific rate filter wrong"
 f!lstMembers.Selected(0) = True: f!cmdApply.Enabled = True: f!txtStamp = "old"
 f!cboRankFilter = "E5": UXDirectoryRankFilter
 Check Not f!cmdApply.Enabled And Len(Nz(f!txtStamp, "")) = 0 And f!lstMembers.ItemsSelected.Count = 0, "Rank change retained stale batch preview/selection"
 f!cboRankFilter = "Unentered": UXDirectoryRankFilter
 Check f!lstMembers.ListCount = 1, "Unentered rank filter wrong"
 DoCmd.Close acForm, "frmUXDirectoryBatch", acSaveNo
 groupID = CreateSummaryGroup("Synthetic E6", "EVAL", "E6")
 DoCmd.OpenForm "frmMajorWorkbench": Forms!frmMajorWorkbench!uxGroupContext = groupID
 Check UXGroupAddReport("frmMajorWorkbench"), "Single member creation picker failed"
 Set f = Forms!frmUXGroupReport
 f!cboRankFilter = "E6": UXGroupMemberFilter
 Check f!cboMember.ListCount = 2, "Single picker paygrade filter wrong"
 f!cboMember = id: UXGroupReportMember
 Check f!txtRank = "IT1", "Selected identity did not populate"
 f!cboRankFilter = "E5": UXGroupMemberFilter
 Check f!cboMember.ListCount = 1 And IsNull(f!cboMember) And Len(Nz(f!txtName, "")) = 0 And IsNull(f!txtMemberVersion), "Hidden member identity survived filter"
 DoCmd.Close acForm, "frmUXGroupReport", acSaveNo
 Check DCount("*", "Report") = 0, "Filtering created reports"
 TestImportRank = "PASS: " & n & " automatic 20xx import and rank-filter assertions": Exit Function
Failed:
 TestImportRank = "FAIL after " & n & ": " & Err.Number & " " & Err.Description
End Function
