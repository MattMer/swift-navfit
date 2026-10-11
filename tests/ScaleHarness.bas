Attribute VB_Name = "ScaleHarness"
Option Explicit
Public Function SeedScaleGroup(ByVal size As Long, ByVal kind As String) As Long
 Dim groupID As Long, i As Long, id As Long, name As String, rate As String, grade As String, designator As Variant, traits As Variant
 grade = IIf(kind = "EVAL", "E5", IIf(kind = "CHIEFEVAL", "E7", "O3"))
 rate = IIf(kind = "EVAL", "IT2", IIf(kind = "CHIEFEVAL", "ITC", "LT"))
 designator = IIf(kind = "FITREP", "1110", Null)
 groupID = CreateSummaryGroup("SYNTHETIC " & kind & " - " & size & " reports", kind, grade)
 For i = 1 To size
  name = "SYNTHETIC " & Format$(groupID, "00") & " MEMBER " & Format$(i, "000")
  traits = Array("4.0", "3.0", "4.0", "4.0", "3.0", "4.0", "4.0")
  If i Mod 5 = 0 Then traits = Array("3.0", "3.0", "3.0", "3.0", "3.0", "3.0", "3.0")
  id = CommitDraft(0, kind, name, name, #1/1/2026#, #9/30/2026#, "Synthetic scale-test report; no real personnel.", "Synthetic scale fixture", "Synthetic performance narrative for member " & i & ". Demonstrated technical competence, teamwork and leadership.", Array(name, rate, designator, "87114", "SYNTHETIC COMMAND"), traits, Array("Observed", "Promotable"), "REGULAR", Array("SYNTHETIC A", "CAPT", "1110", "CO", "87114"))
  AddGroupMember groupID, id, grade
 Next
 SeedScaleGroup = groupID
End Function
Public Function ScaleOpen(ByVal groupID As Long) As Boolean
 Forms!frmUXGroups!lstGroups = groupID
 ScaleOpen = UXOpenGroup()
End Function
Public Function ScaleNavigate(ByVal ordinal As Long) As Long
 Dim f As Form, r As DAO.Recordset
 Set f = Forms!frmUXComparison!uxTraitGrid.Form
 Set r = f.RecordsetClone: r.MoveFirst
 If ordinal > 0 Then r.Move ordinal
 f.Bookmark = r.Bookmark: ScaleNavigate = r!ReportID: r.Close
End Function
Public Function ScaleSave(ByVal ordinal As Long) As Boolean
 Dim id As Long, f As Form, expected As String, revision As Long
 id = ScaleNavigate(ordinal): Set f = Forms!frmUXComparison!uxTraitGrid.Form
 revision = CurrentRevision(id): expected = IIf(Nz(f!Trait33, "") = "4.0", "3.0", "4.0")
 f!Trait33 = expected
 If Not UXGridQueue("Trait33") Then Err.Raise vbObjectError + 8100, , "Queue failed"
 Forms!frmUXComparison.TimerInterval = 0
 If Not UXGridDrain() Then Err.Raise vbObjectError + 8101, , Forms!frmUXComparison!lblStatus.Caption
 If CurrentRevision(id) = revision Or DLookup("Trait33", "Report", "ReportID=" & id) <> expected Then Err.Raise vbObjectError + 8102, , "Scale save did not persist"
 ScaleSave = True
End Function
Public Function ScaleOpenEditor(ByVal ordinal As Long) As Boolean
 Dim id As Long
 id = ScaleNavigate(ordinal)
 DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
 ScaleOpenEditor = True
End Function
Public Function ScaleCloseEditor() As Boolean
 DoCmd.Close acForm, "frmEditor", acSaveNo
 ScaleCloseEditor = Not CurrentProject.AllForms("frmEditor").IsLoaded
End Function
Public Function ScaleHistoryCount() As Long
 ScaleHistoryCount = DCount("*", "GroupSnapshot") + DCount("*", "GroupSnapshotMember") + DCount("*", "ComparisonSet") + DCount("*", "ComparisonMember")
End Function

Public Function ScaleLoadedRows() As Long
    ScaleLoadedRows = Forms!frmUXComparison!uxTraitGrid.Form.Recordset.RecordCount
End Function
