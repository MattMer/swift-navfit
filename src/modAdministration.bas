Attribute VB_Name = "modAdministration"
Option Compare Database
Option Explicit
Public Function AdministrationKeys() As Variant
    AdministrationKeys = Array("OfficialName", "RankRate", "Designator", "UnitUIC", "ShipStation")
End Function
Public Function DraftFieldKeys() As Variant
    DraftFieldKeys = Array("FormKind", "Title", "MemberName", "PeriodStart", "PeriodEnd", "DraftNotes", "OfficialComments", "OfficialName", "RankRate", "Designator", "UnitUIC", "ShipStation", "Trait33", "Trait34", "Trait35", "Trait36", "Trait37", "Trait38", "Trait39")
End Function
Public Function AdministrationLabel(ByVal key As String) As String
    Select Case key
        Case "OfficialName": AdministrationLabel = "Official full name"
        Case "RankRate": AdministrationLabel = "Grade / rate"
        Case "Designator": AdministrationLabel = "Designator / designation"
        Case "UnitUIC": AdministrationLabel = "Unit UIC"
        Case "ShipStation": AdministrationLabel = "Ship / station"
    End Select
End Function
Public Function ValidateAdministration(ByVal values As Variant) As Variant
    Dim i As Long, limits As Variant, keys As Variant
    If Not IsArray(values) Then Err.Raise vbObjectError + 1900, , "Administration must contain five text or null values."
    If LBound(values) <> 0 Or UBound(values) <> 4 Then Err.Raise vbObjectError + 1900, , "Administration must contain five text or null values."
    limits = Array(27, 5, 12, 5, 18): keys = AdministrationKeys()
    For i = 0 To 4
        If Not IsNull(values(i)) Then
            If VarType(values(i)) <> vbString Then Err.Raise vbObjectError + 1901, , AdministrationLabel(keys(i)) & " must be text or null."
            If Len(values(i)) > limits(i) Then Err.Raise vbObjectError + 1902, , AdministrationLabel(keys(i)) & " exceeds " & limits(i) & " characters. Shorten it deliberately; it will not be truncated."
            If InStr(values(i), vbCr) > 0 Or InStr(values(i), vbLf) > 0 Or InStr(values(i), vbTab) > 0 Then Err.Raise vbObjectError + 1903, , AdministrationLabel(keys(i)) & " must be a single line without tabs."
        End If
    Next
    ValidateAdministration = values
End Function
Public Function AdministrationFromEditor() As Variant
    AdministrationFromEditor = Array(Forms!frmEditor!txtOfficialName.Value, Forms!frmEditor!txtRankRate.Value, Forms!frmEditor!txtDesignator.Value, Forms!frmEditor!txtUnitUIC.Value, Forms!frmEditor!txtShipStation.Value)
End Function
Public Function RevisionSnapshotSchema(ByVal revisionID As Long) As Long
    Dim r As DAO.Recordset
    Set r = CurrentDb.OpenRecordset("SELECT SnapshotSchema FROM ReportRevision WHERE RevisionID=" & revisionID, dbOpenSnapshot)
    If Not r.EOF Then RevisionSnapshotSchema = Nz(r!SnapshotSchema, 0)
    r.Close
End Function
Public Function AdministrationFromRevision(ByVal revisionID As Long) As Variant
    Dim values As Variant, keys As Variant, r As DAO.Recordset, i As Long, schema As Long
    values = Array(Null, Null, Null, Null, Null): keys = AdministrationKeys()
    schema = RevisionSnapshotSchema(revisionID)
    For i = 0 To 4
        Set r = CurrentDb.OpenRecordset("SELECT ValueText,IsNullValue FROM RevisionValue WHERE RevisionID=" & revisionID & " AND FieldKey='" & keys(i) & "'", dbOpenSnapshot)
        If r.EOF Then
            If schema >= 9 Then r.Close: Err.Raise vbObjectError + 1904, , "Administrative snapshot is incomplete."
        ElseIf Not r!IsNullValue Then
            If IsNull(r!ValueText) Then r.Close: Err.Raise vbObjectError + 1905, , "Administrative snapshot has an inconsistent null flag."
            values(i) = r!ValueText
        End If
        r.Close
    Next
    AdministrationFromRevision = ValidateAdministration(values)
End Function
