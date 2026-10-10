Attribute VB_Name = "UXDatePickerHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 7785, , message
    checks = checks + 1
End Sub
Public Function TestDatePickers() As String
    Dim kind As Variant, key As Variant, id As Long, revision As Long, f As Access.Form, rows As Access.Form, c As Access.Control
    Dim name As Variant, selected As Variant, checkpoint As Long, values As Object, stage As String
    On Error GoTo Failed
    ConfigureWorkspace "SYNTHETIC CALENDAR TEST", "Tester"
    Check UXDateEntry(#10/6/2026#) = "2026-10-06", "Calendar date not invariant"
    Check UXDateEntry("not-a-date") = "not-a-date", "Incomplete raw date lost"
    Check IsNull(UXDateEntry(Null)), "Null became a date"
    Check UXDateEntry("NOT PERF") = "NOT PERF", "Counseling status lost"
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        id = CommitDraft(0, CStr(kind), "Calendar", "Synthetic", Null, Null, "")
        revision = CurrentRevision(id)
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Set f = Forms!frmEditor: Set rows = f!uxFieldSurface.Form
        For Each key In DraftFieldKeys()
            If UXFieldDate(CStr(key)) And UXFieldApplies(CStr(key), CStr(kind)) Then
                name = UXFieldControl(CStr(key)): stage = kind & " / " & key
                Check UXSelectSection(UXFieldSection(CStr(key))), "Date section inaccessible"
                Set c = rows.Controls(name)
                Check c.ShowDatePicker = 1 And c.Format = "yyyy-mm-dd" And Len(c.InputMask) = 0, "Native calendar not enabled: " & stage
                selected = IIf(key = "PeriodStart", #1/1/2026#, #9/30/2026#)
                c.Value = selected
                Check UXSurfaceEdited(CStr(key)), "Calendar selection callback failed"
                Check f.Controls(name).Value = Format$(CDate(selected), "yyyy-mm-dd"), "Calendar date not normalized: " & stage
                Check CurrentRevision(id) = revision, "Calendar selection saved a revision"
            End If
        Next
        Check CheckpointEditor(), "Date checkpoint failed"
        Set values = EditorPendingFields()
        Check values("DateReported") = "2026-09-30", "Checkpoint date differs"
        Check UXSaveEditor(), "Calendar dates not saveable: " & f!uxError.Caption
        Check DLookup("DateReported", "Report", "ReportID=" & id) = "2026-09-30", "Selected date reverted"
        Check UXSelectSection(3), "Counseling inaccessible"
        rows!txtCounselingDate = "NOT PERF": UXSurfaceEdited "CounselingDate"
        Check UXSaveEditor(), "Calendar formatting rejected counseling status"
        Check DLookup("CounselingDate", "Report", "ReportID=" & id) = "NOT PERF", "Status did not save"
        Check UXSelectSection(0), "Admin inaccessible"
        rows!txtDateReported = "not-a-date": UXSurfaceEdited "DateReported"
        Check Not UXSaveEditor(), "Invalid typed date saved"
        Check CheckpointEditor(), "Incomplete date not checkpointed"
        checkpoint = DLookup("RecoveryDraftID", "RecoveryDraft", "ReportID=" & id)
        Check DLookup("DateReported", "RecoveryDraft", "RecoveryDraftID=" & checkpoint) = "not-a-date", "Calendar lost raw incomplete entry"
        UXLeaveEditor "Discard"
    Next
    For Each name In Array("frmUXSharedFields", "frmUXPriorEditor", "frmUXSeniorEditor", "frmMyWork")
        DoCmd.OpenForm CStr(name), acDesign
        Set f = Forms(CStr(name))
        For Each c In f.Controls
            If c.Name = "valuePeriodStart" Or c.Name = "valuePeriodEnd" Or c.Name = "valueCounselingDate" Or c.Name = "txtFrom" Or c.Name = "txtTo" Or c.Name = "txtAsOf" Or c.Name = "txtDue" Then Check c.ShowDatePicker = 1, "Context date calendar missing"
        Next
        DoCmd.Close acForm, CStr(name), acSaveNo
    Next
    TestDatePickers = "PASS: " & checks & " native calendar, canonical dates, counseling status and raw recovery assertions": Exit Function
Failed:
    TestDatePickers = "FAIL after " & checks & " at " & stage & ": " & Err.Number & " " & Err.Description
End Function

