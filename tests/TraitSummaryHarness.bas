Attribute VB_Name = "TraitSummaryHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 2100, , message
    checks = checks + 1
End Sub
Public Function TestTraitSummary() As String
    Dim values As Variant, text As String, kind As Variant, id As Long, revisions As Long, checkpoints As Long
    On Error GoTo Failed
    values = Array(Null, "", Null, "", Null, "", Null)
    text = TraitSummary(values)
    Check InStr(text, "no numeric marks") > 0 And InStr(text, "blank 7") > 0, "Blank arithmetic wrong"
    values = Array("NOB", "NOB", "NOB", "NOB", "NOB", "NOB", "NOB")
    text = TraitSummary(values)
    Check InStr(text, "no numeric marks") > 0 And InStr(text, "NOB 7") > 0, "All NOB became zero"
    values = Array("1.0", "2.0", "3.0", "4.0", "5.0", "NOB", Null)
    text = TraitSummary(values)
    Check InStr(text, Format$(3, "0.00")) > 0 And InStr(text, "numeric 5; NOB 1; blank 1") > 0, "NOB/blank denominator wrong"
    Check InStr(text, "partial") > 0, "Partial average unlabelled"
    values = Array("1.0", "1.0", "2.0", "NOB", "NOB", "NOB", "NOB")
    text = TraitSummary(values)
    Check InStr(text, Format$(4 / 3, "0.00")) > 0 And InStr(text, "partial") = 0, "Rounding or NOB completeness wrong"
    values = Array("5.0", "5.0", "5.0", "5.0", "5.0", "5.0", "5.0")
    Check InStr(TraitSummary(values), Format$(5, "0.00")) > 0, "Upper boundary wrong"
    values(0) = "4.5"
    Check InStr(TraitSummary(values), "unsupported") > 0, "Invalid grade averaged"
    values(0) = 4
    Check InStr(TraitSummary(values), "unsupported") > 0, "Numeric code averaged"
    Check InStr(TraitSummary(Array("1.0")), "unsupported") > 0, "Malformed array averaged"
    Call ConfigureWorkspace("Summary test", "Synthetic Tester")
    DoCmd.OpenForm "frmEditor"
    Check Forms!frmEditor!txtSavedTraitSummary = "Saved values: no committed revision.", "New editor falsely saved"
    Check Forms!frmEditor!txtPendingTraitSummary.Locked And Forms!frmEditor!txtSavedTraitSummary.Locked, "Derived controls editable"
    Call FixtureCloseEditor()
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        values = Array("1.0", "2.0", "3.0", "4.0", "5.0", "NOB", Null)
        id = CommitDraft(0, CStr(kind), "Summary test", "Synthetic", Null, Null, "", "Test", "", , values)
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Check InStr(Forms!frmEditor!txtSavedTraitSummary, "Saved revision 1") > 0, "Saved revision unidentified"
        Check InStr(Forms!frmEditor!txtSavedTraitSummary, Format$(3, "0.00")) > 0, "Saved average wrong"
        revisions = DCount("*", "ReportRevision"): checkpoints = DCount("*", "RecoveryDraft")
        Forms!frmEditor!cboTrait33 = "5.0"
        Call ShowTraitGuidance(33)
        Check InStr(Forms!frmEditor!txtPendingTraitSummary, Format$(3.8, "0.00")) > 0, "Editor average not refreshed"
        Check InStr(Forms!frmEditor!txtSavedTraitSummary, Format$(3, "0.00")) > 0, "Pending edit changed saved average"
        Check DCount("*", "ReportRevision") = revisions And DCount("*", "RecoveryDraft") = checkpoints, "Summary wrote history/checkpoint"
        Check SaveEditor(), "Summary editor save failed"
        Check InStr(Forms!frmEditor!txtSavedTraitSummary, "Saved revision 2") > 0 And InStr(Forms!frmEditor!txtSavedTraitSummary, Format$(3.8, "0.00")) > 0, "Save did not refresh summary"
        Forms!frmEditor!cboTrait33 = "INVALID"
        Call RefreshTraitSummary()
        Check InStr(Forms!frmEditor!txtPendingTraitSummary, "unsupported") > 0, "Invalid editor value averaged"
        Check InStr(Forms!frmEditor!txtSavedTraitSummary, Format$(3.8, "0.00")) > 0, "Invalid input corrupted saved summary"
        Call CheckpointEditor()
        Call FixtureCloseEditor()
        Call RestoreRecoveryDraft(DMax("RecoveryDraftID", "RecoveryDraft"))
        Check InStr(Forms!frmEditor!txtPendingTraitSummary, "unsupported") > 0, "Recovered summary stale"
        Check InStr(Forms!frmEditor!txtSavedTraitSummary, Format$(3.8, "0.00")) > 0, "Recovery changed saved summary"
        Call FixtureCloseEditor()
    Next
    TestTraitSummary = "PASS: " & checks & " trait summary assertions": Exit Function
Failed:
    TestTraitSummary = "FAIL: " & Err.Number & " " & Err.Description
End Function

Private Sub FixtureCloseEditor()
    ' Preserve the old abrupt-close semantics for persistence/recovery fixtures.
    ' Real save/discard/cancel exits are exercised by UXWorkflowHarness.
    Forms!frmEditor.OnUnload = ""
    DoCmd.Close acForm, "frmEditor", acSaveNo
End Sub
