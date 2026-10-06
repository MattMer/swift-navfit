Attribute VB_Name = "SmokeHarness"
Option Compare Database
Option Explicit
Public Function SmokeFormFits(ByVal name As String) As Boolean
    Dim f As Access.Form, c As Access.Control
    Set f = Forms(name)
    For Each c In f.Controls
        If c.Visible And (c.Left < 0 Or c.Left + c.Width > f.Width) Then Exit Function
    Next
    SmokeFormFits = True
End Function
Public Function SmokeEditorEntries() As Boolean
    Forms!frmEditor!cboKind = "CHIEFEVAL"
    Forms!frmEditor!txtTitle = "Synthetic UI draft"
    Forms!frmEditor!txtNotes = "Saved through editor action"
    SmokeEditorEntries = True
End Function
Public Function RejectArchived(ByVal reportID As Long) As Boolean
    On Error GoTo Rejected
    CommitDraft reportID, "EVAL", "Forbidden", "Synthetic Member", Null, Null, "Forbidden", "Archive rejection"
    Exit Function
Rejected:
    RejectArchived = (Err.Number = vbObjectError + 110)
End Function

Public Function RejectFolderCycle(ByVal folderID As Long, ByVal descendantID As Long) As Boolean
    On Error GoTo Rejected
    MoveFolder folderID, descendantID
    Exit Function
Rejected:
    RejectFolderCycle = (Err.Number = vbObjectError + 205)
End Function

Public Function RejectTypeChange(ByVal reportID As Long) As Boolean
    On Error GoTo Rejected
    CommitDraft reportID, "FITREP", "Forbidden", "Synthetic Member", Null, Null, "Forbidden", "Type rejection"
    Exit Function
Rejected:
    RejectTypeChange = (Err.Number = vbObjectError + 112)
End Function

Public Function RejectSnapshotFailure(ByVal reportID As Long) As Boolean
    Dim db As DAO.Database, rule As String
    Set db = CurrentDb
    rule = db.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule
    db.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = "<> 'DraftNotes'"
    On Error GoTo Rejected
    CommitDraft reportID, "EVAL", "MUST ROLL BACK", "Synthetic Member", Null, Null, "Injected failure", "Failure injection"
    db.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = rule
    Exit Function
Rejected:
    RejectSnapshotFailure = True
    db.TableDefs("RevisionValue").Fields("FieldKey").ValidationRule = rule
End Function

Public Function SmokeAddGroupMember(ByVal groupID As Long, ByVal reportID As Long, ByVal paygrade As String, ByVal reason As String) As Boolean
    Call AddGroupMember(groupID, reportID, paygrade, reason)
    SmokeAddGroupMember = True
End Function
