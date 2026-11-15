Attribute VB_Name = "SnapshotFixtures"
Option Explicit
' Test fixtures for snapshot storage (L03). From schema 50 a save stores only fields
' with a value, so a missing row means an empty field. Fixtures that model an older
' full-copy revision, or damage one, first turn the revision into a full copy: every
' catalogue field it holds gets a row (empty ones flagged) and ValueEncoding is cleared.
Public Sub SnapshotFullCopy(ByVal revisionID As Long)
    Dim key As Variant, db As DAO.Database
    Set db = CurrentDb
    If Nz(DLookup("ValueEncoding", "ReportRevision", "RevisionID=" & revisionID), 0) <> 1 Then Exit Sub
    For Each key In DraftFieldKeys()
        If SparseHoldsField(revisionID, CStr(key)) Then
            If DCount("*", "RevisionValue", "RevisionID=" & revisionID & " AND FieldKey=" & QuoteSQL(CStr(key))) = 0 Then
                db.Execute "INSERT INTO RevisionValue (RevisionID,FieldKey,ValueText,IsNullValue) VALUES (" & revisionID & "," & QuoteSQL(CStr(key)) & ",Null,True)", dbFailOnError
            End If
        End If
    Next
    db.Execute "UPDATE ReportRevision SET ValueEncoding=Null WHERE RevisionID=" & revisionID, dbFailOnError
End Sub
' Every catalogue field a revision holds is present: a row, or empty in a sparse revision.
Public Function SnapshotComplete(ByVal revisionID As Long) As Boolean
    Dim key As Variant
    For Each key In DraftFieldKeys()
        If Not MajorRevisionHasField(revisionID, CStr(key)) Then Exit Function
    Next
    SnapshotComplete = True
End Function
