Attribute VB_Name = "ProvenanceHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 1650, , message
    checks = checks + 1
End Sub
Private Function InvalidRowRejected() As Boolean
    On Error GoTo Rejected
    Call LoadProvenanceFields()
    Exit Function
Rejected:
    InvalidRowRejected = True
End Function
Private Function MissingReceiptRejected() As Boolean
    On Error GoTo Rejected
    Call OpenImportProvenance(0)
    Exit Function
Rejected:
    MissingReceiptRejected = True
End Function
Public Function TestProvenance() As String
    Dim receipt As Long, r As DAO.Recordset, id As Long, before As Long, revisions As Long, count As Long, text As String
    On Error GoTo Failed
    receipt = DMin("ImportReceiptID", "ImportReceipt")
    Set r = CurrentDb.OpenRecordset("ImportRawValue", dbOpenDynaset)
    For count = 1 To 3
        r.AddNew: r!ImportReceiptID = receipt: r!SourceTable = "Reports": r!SourceRow = 999
        r!FieldOrdinal = count: r!FieldKey = "Test'Field" & count: r!DAOType = 12: r!IsNullValue = (count = 1)
        If count = 2 Then r!RawValue = ""
        If count = 3 Then r!RawValue = String$(9000, "Q") & vbCrLf & "Case'End"
        r.Update
    Next
    r.Close
    before = DCount("*", "ImportRawValue"): revisions = DCount("*", "ReportRevision")
    Check MissingReceiptRejected(), "Missing receipt accepted"
    Check OpenImportProvenance(receipt), "Viewer failed"
    Check Forms!frmProvenance!txtValue.Locked And Forms!frmProvenance!txtContext.Locked, "Viewer editable"
    Check InStr(Forms!frmProvenance!txtContext, DLookup("SourceHash", "ImportReceipt", "ImportReceiptID=" & receipt)) > 0, "Hash absent"
    Check InStr(Forms!frmProvenance!txtContext, "encoding:") > 0, "Encoding absent"
    Check Forms!frmProvenance!lstFields.ListCount = DCount("*", "ImportRawValue", "ImportReceiptID=" & receipt & " AND SourceTable='Reports'"), "Receipt scope wrong"
    Forms!frmProvenance!txtRow = 999: Forms!frmProvenance!txtSearch = "Test'Field"
    Check LoadProvenanceFields(), "Filter failed"
    Check Forms!frmProvenance!lstFields.ListCount = 3, "Apostrophe search failed"
    For count = 1 To 3
        id = DLookup("ImportRawValueID", "ImportRawValue", "ImportReceiptID=" & receipt & " AND SourceRow=999 AND FieldOrdinal=" & count)
        Forms!frmProvenance!lstFields = id
        Check ShowProvenanceValue(), "Selection failed"
        Select Case count
            Case 1: Check InStr(Forms!frmProvenance!lblValue.Caption, "NULL") > 0, "Null marker absent"
            Case 2: Check InStr(Forms!frmProvenance!lblValue.Caption, "Empty text") > 0, "Empty marker absent"
            Case 3: Check StrComp(Forms!frmProvenance!txtValue, String$(9000, "Q") & vbCrLf & "Case'End", vbBinaryCompare) = 0, "Long text changed"
        End Select
    Next
    Forms!frmProvenance!txtRow = 0: Forms!frmProvenance!txtSearch = ""
    Check LoadProvenanceFields(), "Zero row rejected"
    Check Forms!frmProvenance!lstFields.ListCount = 128, "Zero row missing"
    Forms!frmProvenance!txtRow = -1: Check InvalidRowRejected(), "Negative row accepted"
    Forms!frmProvenance!txtRow = "1 OR 1=1": Check InvalidRowRejected(), "Invalid row accepted"
    Forms!frmProvenance!txtRow = 0.5: Check InvalidRowRejected(), "Fractional row accepted"
    Forms!frmProvenance!txtRow = Null: Forms!frmProvenance!cboTable = "Summary"
    Check LoadProvenanceFields(), "Empty table failed"
    Check Forms!frmProvenance!lstFields.ListCount = 0 And IsNull(Forms!frmProvenance!txtValue), "Stale value retained"
    Forms!frmProvenance!cboTable = "All": Check LoadProvenanceFields(), "All filter failed"
    Check Forms!frmProvenance!lstFields.ListCount = DCount("*", "ImportRawValue", "ImportReceiptID=" & receipt), "All fields count wrong"
    Check DCount("*", "ImportRawValue") = before And DCount("*", "ReportRevision") = revisions, "Viewer changed saved data"
    DoCmd.Close acForm, "frmProvenance", acSaveNo
    TestProvenance = "PASS: " & checks & " provenance assertions": Exit Function
Failed:
    TestProvenance = "FAIL: " & Err.Number & " " & Err.Description
End Function
