Attribute VB_Name = "PDFExpandedHarness"
Option Compare Binary
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 4750, , message
    checks = checks + 1
End Sub
Private Function RejectedApply(ByVal target As Long, ByVal mode As String) As Boolean
    Dim id As Long
    On Error Resume Next
    id = ImportPDFPreview("20", "SSN", target, mode)
    RejectedApply = Err.Number <> 0: Err.Clear
    On Error GoTo 0
End Function
Public Function TestExpandedPDF(ByVal directory As String) As String
    Dim kind As Variant, p As Object, values As Object, selected As Object, old As Object, changes As Object
    Dim id As Long, revision As Long, before As Long, receipt As Long, target As Long, rejected As Boolean, key As Variant, path As String
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC PDF EXPANSION", "Synthetic PDF tester"
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        path = directory & "\" & LCase$(kind) & "-expanded.pdf"
        Check Left$(PDFExtractionStatus(path), 5) = "PASS:", "Dependency cannot extract " & kind
        before = DCount("*", "ReportRevision")
        Set p = PreviewStructuredPDF(path): Set values = PDFMappedValues("20", "SSN")
        Check values("CommandEmployment") = "SYNTHETIC BLOCK 28: Training and command achievements.", "Block 28 text lost: " & kind
        Check values.Count = PDFMappingCount(CStr(kind)), "Expanded field count wrong: " & kind & " / " & values.Count
        Check DCount("*", "ReportRevision") = before, "Extraction/preview wrote a revision"
        Check values("UnitUIC") = "12345" And values("ReportingSeniorUIC") = "54321", "UIC fields swapped"
        Check values("ReportingSeniorSSN") = "000000009", "Reporting senior SSN lost"
        Check values("PrimaryDuty") = "ADMIN" And values("DutyDescription") = "O'Brien & synthetic duties.", "Duty fields swapped"
        Check values("CounselingDate") = "2026-01-01" And values("DateReported") = "2026-01-01", "Date decoding wrong"
        Check values("PromotionStatus") = "REGULAR" And values("PromotionRecommendation") = "Must Promote", "Option code decoding wrong"
        Check values("ComponentActive") = "Yes" And values("ComponentTAR") = "No", "Component decoding wrong"
        For Each key In TraitKeys()
            Check values(key) = CStr(3 + (CLng(Right$(key, 2)) Mod 3)) & ".0", "Trait decoding wrong: " & key
        Next
        If kind = "CHIEFEVAL" Then
            Check values("OccasionPeriodic") = "Yes" And values("ReportConcurrent") = "No", "Chief checkboxes wrong"
            Check values("ReportingSeniorAddress1") = "123 TEST ROAD" And values("ReportingSeniorPhone") = "0000000000", "Chief contacts wrong"
        Else
            Check Not values.Exists("OccasionPeriodic"), "Unproven unnamed checkbox imported"
        End If
        id = ImportPDFPreview("20", "SSN")
        Set old = SavedFields(id)
        For Each key In values.Keys
            Check SameFieldValue(old(key), values(key)), "Saved expanded field differs: " & key
        Next
        Set p = PreviewStructuredPDF(directory & "\" & LCase$(kind) & "-merge.pdf")
        target = CommitDraft(0, CStr(kind), "Existing target", "Preserved display name", Null, Null, "Preserve local notes", "PDF merge fixture")
        Set changes = CreateObject("Scripting.Dictionary"): changes.Add "OfficialName", "PRESERVED NAME": changes.Add "MemberFirstName", "Separate first": changes.Add "OfficialComments", "PREVIOUS COMMENTS"
        Call CommitSavedFields(target, CurrentRevision(target), changes, "Prepare existing values")
        Set old = SavedFields(target): revision = CurrentRevision(target): before = DCount("*", "Report")
        Set selected = PDFPrepareMerge("20", "SSN", target, "Fill empty fields")
        Check Not selected.Exists("OfficialName") And Not selected.Exists("OfficialComments") And selected.Exists("UnitUIC"), "Fill-empty mode overwrites existing fields"
        Check CurrentRevision(target) = revision, "Merge preview created revision"
        Set changes = CreateObject("Scripting.Dictionary"): changes.Add "ShipStation", "EDIT AFTER PREVIEW"
        Call CommitSavedFields(target, revision, changes, "Concurrent edit")
        revision = CurrentRevision(target)
        Check RejectedApply(target, "Fill empty fields") And CurrentRevision(target) = revision, "Stale merge changed target"
        Set selected = PDFPrepareMerge("20", "SSN", target, "Replace mapped values")
        Check selected.Exists("OfficialComments") And selected.Exists("OfficialName"), "Replacement mode does not show conflicts"
        Check RejectedApply(target, "Fill empty fields"), "Changed merge mode accepted without preview"
        id = ImportPDFPreview("20", "SSN", target, "Replace mapped values")
        Check id = target And DCount("*", "Report") = before, "Merge created duplicate report"
        Check MajorScalar("SELECT OfficialComments FROM Report WHERE ReportID=" & target) = "MERGED SYNTHETIC COMMENTS", "Merge did not replace comments"
        Check MajorScalar("SELECT MemberFirstName FROM Report WHERE ReportID=" & target) = "Separate first", "Structured name overwritten"
        Check MajorScalar("SELECT MemberName FROM Report WHERE ReportID=" & target) = "Preserved display name", "Display name overwritten"
        Check MajorScalar("SELECT DraftNotes FROM Report WHERE ReportID=" & target) = "Preserve local notes", "Local notes overwritten"
        receipt = Nz(DMax("ImportReceiptID", "ImportReceipt", "ReportID=" & target), 0)
        Check MajorScalar("SELECT BeforeRevisionID FROM ImportReceipt WHERE ImportReceiptID=" & receipt) = revision, "Merge baseline provenance absent"
        Check UndoNarrativeImport(receipt), "Existing guarded import undo failed"
        Check MajorScalar("SELECT OfficialComments FROM Report WHERE ReportID=" & target) = "PREVIOUS COMMENTS", "Undo did not restore old comments"
        Check MajorScalar("SELECT ShipStation FROM Report WHERE ReportID=" & target) = "EDIT AFTER PREVIEW", "Undo lost preceding edits"
        Check Not CBool(MajorScalar("SELECT IsArchived FROM Report WHERE ReportID=" & target)), "Merge undo archived existing draft"
        Set p = PreviewStructuredPDF(path)
        Check RejectedApply(0, "Fill empty fields"), "Cross-adapter duplicate guard failed"
        ' Exercise conversion failures and omitted values without touching source bytes.
        Set p = PreviewStructuredPDF(path)
        If kind = "CHIEFEVAL" Then key = "TECHDROPDOWN" Else key = "_5"
        p(key) = "9"
        revision = DCount("*", "ReportRevision"): before = DCount("*", "ImportReceipt")
        On Error Resume Next
        Set values = PDFMappedValues("20", "SSN"): rejected = Err.Number <> 0: Err.Clear
        On Error GoTo Failed
        Check rejected And DCount("*", "ReportRevision") = revision And DCount("*", "ImportReceipt") = before, "Malformed trait wrote state"
        p(key) = "": Set values = PDFMappedValues("20", "SSN")
        Check Not values.Exists("Trait33"), "Blank trait became a destructive clear"
        p(key) = Null: Set values = PDFMappedValues("20", "SSN")
        Check Not values.Exists("Trait33"), "Missing trait became a destructive clear"
        p(key) = "NOB": If kind = "CHIEFEVAL" Then p(key) = "0"
        Set values = PDFMappedValues("20", "SSN")
        Check values("Trait33") = "NOB", "NOB code lost"
        Call ArchiveDraft(target, True)
        On Error Resume Next
        Set selected = PDFPrepareMerge("20", "SSN", target, "Replace mapped values"): rejected = Err.Number <> 0: Err.Clear
        On Error GoTo Failed
        Check rejected, "Archived merge target accepted"
        Call ArchiveDraft(target, False)
        CurrentDb.Execute "UPDATE Report SET WorkflowState='Approved' WHERE ReportID=" & target, dbFailOnError
        On Error Resume Next
        Set selected = PDFPrepareMerge("20", "SSN", target, "Replace mapped values"): rejected = Err.Number <> 0: Err.Clear
        On Error GoTo Failed
        Check rejected, "Locked merge target accepted"
        CurrentDb.Execute "UPDATE Report SET WorkflowState='Draft' WHERE ReportID=" & target, dbFailOnError
        CurrentDb.Execute "INSERT INTO RecoveryDraft (SessionGuid,AuthorProfileID,ReportID,BaseRevisionID) VALUES (" & QuoteSQL(NewIdentity()) & "," & MajorAuthor() & "," & target & "," & CurrentRevision(target) & ")", dbFailOnError
        On Error Resume Next
        Set selected = PDFPrepareMerge("20", "SSN", target, "Replace mapped values"): rejected = Err.Number <> 0: Err.Clear
        On Error GoTo Failed
        Check rejected, "Target with pending recovery accepted"
        CurrentDb.Execute "DELETE FROM RecoveryDraft WHERE ReportID=" & target, dbFailOnError
        If kind <> "FITREP" Then
            revision = CommitDraft(0, "FITREP", "Wrong type", "Synthetic wrong type", Null, Null, "", "Type mismatch fixture")
            On Error Resume Next
            Set selected = PDFPrepareMerge("20", "SSN", revision, "Replace mapped values"): rejected = Err.Number <> 0: Err.Clear
            On Error GoTo Failed
            Check rejected, "Different-form merge target accepted"
        End If
        Call PDFClearPreview
    Next
    TestExpandedPDF = "PASS: " & checks & " expanded extraction/mapping/merge/undo assertions": Exit Function
Failed:
    TestExpandedPDF = "FAIL after " & checks & ": " & Err.Number & " / " & Err.Description
End Function
