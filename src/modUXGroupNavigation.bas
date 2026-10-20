Attribute VB_Name = "modUXGroupNavigation"
Option Explicit
'One exit boundary for group pages. Invalid queued cells keep the existing
'group open; successful exits dispose the cache through the native unload guard.
Public Function UXLeaveGroup() As Boolean
 Dim name As Variant, f As Access.Form
 On Error GoTo Failed
 If CurrentProject.AllForms("frmUXComparison").IsLoaded Then
  If Not UXGridDrain() Then Exit Function
 End If
 If CurrentProject.AllForms("frmUXBulk").IsLoaded Then
  Set f = Forms!frmUXBulk
  f!cmdBack.SetFocus
  If UXGroupBulkPending() Then Exit Function
  If Not UXCloseBlocks() Then Exit Function
  If CurrentProject.AllForms("frmUXBulk").IsLoaded Then Exit Function
 End If
 If CurrentProject.AllForms("frmUXComparison").IsLoaded Then
  If Not UXCloseComparison(False) Then Exit Function
 End If
 For Each name In Array("frmUXGroupRSCA", "frmMajorWorkbench")
  If CurrentProject.AllForms(CStr(name)).IsLoaded Then
   DoCmd.Close acForm, CStr(name), acSaveNo
   If CurrentProject.AllForms(CStr(name)).IsLoaded Then Exit Function
  End If
 Next
 UXLeaveGroup = True
 Exit Function
Failed:
 If CurrentProject.AllForms("frmUXComparison").IsLoaded Then Forms!frmUXComparison!lblStatus.Caption = "Unable to leave group: " & Err.Description
 If CurrentProject.AllForms("frmUXBulk").IsLoaded Then Forms!frmUXBulk!uxStatus.Caption = "Unable to leave group: " & Err.Description
End Function
Public Function UXBackToGroups() As Boolean
 UXBackToGroups = UXNavigate("frmUXGroups")
End Function
