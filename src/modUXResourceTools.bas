Attribute VB_Name = "modUXResourceTools"
Option Explicit
' Live group tools pass explicit IDs and flush queued comparison edits first.
Public Function UXCurrentGroupTool(ByVal screen As String, ByVal action As String) As Boolean
 Dim id As Long
 On Error GoTo Failed
 id = Nz(Forms(screen)!uxGroupContext, 0): If id = 0 Then Exit Function
 If Not UXGridDrain() Then Exit Function
 Select Case action
  Case "Export": UXCurrentGroupTool = UXOpenNAVFITGroupExport(id)
  Case "Senior": UXCurrentGroupTool = RSOpenGroupCalculator(id)
 End Select
 Exit Function
Failed:
 If UXHasControl(Forms(screen), "lblStatus") Then
  Forms(screen)!lblStatus.Caption = Err.Description
 ElseIf UXHasControl(Forms(screen), "uxStatus") Then
  Forms(screen)!uxStatus.Caption = Err.Description
 End If
End Function
Public Function UXGroupSeniorCaption(ByVal id As Long, Optional ByVal preparedPSG As Object = Nothing) As String
 Dim title As String, db As DAO.Database, r As DAO.Recordset, psg As Object
 title = Nz(MajorScalar("SELECT s.Title FROM GroupRSCABaseline AS g INNER JOIN ReportingSeniorRecord AS s ON g.SeniorID=s.SeniorID WHERE g.GroupID=" & id), "")
 If title = "" Then
  Set db = CurrentDb: Set r = db.OpenRecordset("SELECT DISTINCT r.ReportingSeniorName FROM Report AS r INNER JOIN GroupMember AS m ON r.ReportID=m.ReportID WHERE m.GroupID=" & id, dbOpenSnapshot)
  If Not r.EOF Then
   title = Nz(r!ReportingSeniorName, "Not set"): If title = "" Then title = "Not set"
   r.MoveNext: If Not r.EOF Then title = "Mixed report values"
  End If
  r.Close
 End If
 If title = "" Then title = "Not selected"
 If preparedPSG Is Nothing Then Set psg = RSGroupPSG(id) Else Set psg = preparedPSG
 UXGroupSeniorCaption = "Reporting senior: " & title & " | PSG RSCA: "
 If IsNull(psg("Value")) Then
  UXGroupSeniorCaption = UXGroupSeniorCaption & "Unavailable: " & psg("Reason")
 Else
  UXGroupSeniorCaption = UXGroupSeniorCaption & Format$(psg("Value"), "0.00") & IIf(psg("Estimated"), " (estimate)", "") & IIf(psg("Partial"), " (partial traits)", "")
 End If
End Function
Public Sub UXRefreshGroupSenior()
 If CurrentProject.AllForms("frmUXComparison").IsLoaded Then
  If UXHasControl(Forms!frmUXComparison, "uxSeniorCaption") Then Forms!frmUXComparison!uxSeniorCaption.Caption = UXGroupSeniorCaption(Nz(Forms!frmUXComparison!uxGroupContext, 0))
 End If
End Sub
' Resources opens a dedicated template view, without a summary-group workbench.
Public Function UXOpenRouteTemplates() As Boolean
 If Not UXLeaveEditor() Then Exit Function
 DoCmd.OpenForm "frmUXRouteTemplates"
 UXOpenRouteTemplates = True
End Function
Public Function UXLoadRouteTemplates() As Boolean
 With Forms!frmUXRouteTemplates
  !lstTemplates.RowSource = "SELECT TemplateID,Title FROM RouteTemplate WHERE IsArchived=False ORDER BY Title"
  !cboReviewer.RowSource = "SELECT UserProfileID,DisplayName FROM UserProfile ORDER BY DisplayName,UserProfileID"
 End With
 UXLoadRouteTemplates = True
End Function
Public Function UXShowTemplate() As Boolean
 With Forms!frmUXRouteTemplates
  !lstSteps.RowSource = "SELECT s.StepOrdinal AS [Step],u.DisplayName AS Contributor,s.Instructions FROM RouteTemplateStep AS s INNER JOIN UserProfile AS u ON s.AssigneeID=u.UserProfileID WHERE s.TemplateID=" & Nz(!lstTemplates, 0) & " ORDER BY s.StepOrdinal"
 End With
 UXShowTemplate = True
End Function
Public Function UXTemplateReviewer(ByVal clear As Boolean) As Boolean
 Dim f As Access.Form
 Set f = Forms!frmUXRouteTemplates
 If clear Then
  f!txtIDs = Null: f!txtReviewers = Null
 ElseIf Nz(f!cboReviewer, 0) > 0 Then
  f!txtIDs = Nz(f!txtIDs, "") & IIf(Len(Nz(f!txtIDs, "")) > 0, ",", "") & f!cboReviewer
  f!txtReviewers = Nz(f!txtReviewers, "") & IIf(Len(Nz(f!txtReviewers, "")) > 0, " > ", "") & f!cboReviewer.Column(1)
 End If
 UXTemplateReviewer = True
End Function
Public Function UXSaveTemplate() As Boolean
 Dim id As Long
 On Error GoTo Failed
 With Forms!frmUXRouteTemplates
  id = CreateRouteTemplate(Nz(!txtTitle, ""), Nz(!txtIDs, ""), Nz(!txtInstructions, ""))
  !lstTemplates.Requery: !lstTemplates = id: UXShowTemplate
  !txtTitle = Null: !txtIDs = Null: !txtReviewers = Null: !txtInstructions = Null
  !lblStatus.Caption = "Template saved. Use it from a report's Review section."
 End With
 UXSaveTemplate = True: Exit Function
Failed:
 Forms!frmUXRouteTemplates!lblStatus.Caption = Err.Description & " Entries retained."
End Function
Public Function UXCloseTemplates() As Boolean
 DoCmd.Close acForm, "frmUXRouteTemplates", acSaveNo
 UXCloseTemplates = UXNavigate("frmUXResources")
End Function

' Inspect the selected reusable senior record directly from its group calculator.
Public Function RSViewGroupSenior() As Boolean
 Dim id As Long
 id = Nz(Forms!frmUXGroupRSCA!cboSenior, 0)
 If id = 0 Then Forms!frmUXGroupRSCA!lblStatus.Caption = "Choose a reporting senior record first.": Exit Function
 If Not RSLeaveSeniorEditor() Then Exit Function
 DoCmd.OpenForm "frmUXSeniorEditor", , , , , , CStr(id)
 RSViewGroupSenior = True
End Function
