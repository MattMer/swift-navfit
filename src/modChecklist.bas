Attribute VB_Name = "modChecklist"
Option Explicit
Public Function OpenDraftChecklist(ByVal reportID As Long) As Boolean
 OpenDraftChecklist = UXOpenReportTool(reportID, "Validation")
End Function
Public Function ChecklistFromEditor() As Boolean
    On Error GoTo Failed
    ChecklistFromEditor = OpenDraftChecklist(Nz(Forms!frmEditor!txtReportID, 0))
    Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Saved draft checklist"
End Function
Public Function ChecklistFromLibrary() As Boolean
    On Error GoTo Failed
    ChecklistFromLibrary = OpenDraftChecklist(Nz(Forms!frmLibrary!lstReports, 0))
    Exit Function
Failed:
    MsgBox Err.Description, vbInformation, "Saved draft checklist"
End Function
Public Function ChecklistField(ByVal key As String) As String
    Select Case key
        Case "TITLE": ChecklistField = "Title"
        Case "MEMBER": ChecklistField = "MemberName"
        Case "START", "DATE_ORDER": ChecklistField = "PeriodStart"
        Case "END": ChecklistField = "PeriodEnd"
        Case "NARRATIVE": ChecklistField = "OfficialComments"
        Case "TRAIT_NOB": ChecklistField = "Trait33"
        Case "POLICY", "EXPORT", "PARTIAL_IMPORT", "RECOVERY", "ARCHIVED": ChecklistField = ""
        Case Else: ChecklistField = key
    End Select
End Function
Public Function RefreshDraftChecklist() As Boolean
 RefreshDraftChecklist = UXToolsValidation()
End Function
Public Function ChecklistGuidance() As Boolean
 ChecklistGuidance = UXToolsFinding()
End Function
Public Function ChecklistOpenSection() As Boolean
 ChecklistOpenSection = UXToolsFindingOpen()
End Function
Public Function CloseDraftChecklist() As Boolean
 CloseDraftChecklist = UXSelectSection(0)
End Function