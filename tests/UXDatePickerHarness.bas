Attribute VB_Name = "UXDatePickerHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 7785, , message
    checks = checks + 1
End Sub
Public Function TestDatePickers() As String
    Dim kind As Variant, key As Variant, id As Long, revision As Long, f As Access.Form, rows As Access.Form, c As Access.Control
    Dim name As Variant, selected As Variant, checkpoint As Long, values As Object, stage As String, status As Variant
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
                If key = "CounselingDate" Or key = "ReadinessDate" Then
                    Check c.Format = "@" And c.ShowDatePicker = 0 And Len(c.InputMask) = 0, key & " text properties: format=" & c.Format & "; picker=" & c.ShowDatePicker & "; mask=" & c.InputMask
                Else
                    Check c.ShowDatePicker = 1 And c.Format = "yyyy-mm-dd" And Len(c.InputMask) = 0, "Native calendar not enabled: " & stage
                End If
                selected = IIf(key = "PeriodStart", #1/1/2026#, #9/30/2026#)
                c.Value = selected
                Check UXSurfaceEdited(CStr(key)), "Calendar selection callback failed"
                If key = "ReadinessDate" Then
                    Check f.Controls(name).Value = Format$(DateSerial(Year(selected), Month(selected), 1), "yyyy-mm-dd") And c.Value = Format$(selected, "mm/yy"), "PFA date not kept as its month: " & f.Controls(name).Value & " / " & c.Value
                Else
                    Check f.Controls(name).Value = Format$(CDate(selected), "yyyy-mm-dd"), "Calendar date not normalized: " & stage
                End If
                Check CurrentRevision(id) = revision, "Calendar selection saved a revision"
            End If
        Next
        ' Block 20: only the Chief form has a PFA date, entered and shown as MM/YY.
        Check UXFieldApplies("ReadinessDate", CStr(kind)) = (kind = "CHIEFEVAL"), "PFA date applies to " & kind
        If kind = "CHIEFEVAL" Then
            Check UXSelectSection(UXFieldSection("ReadinessDate")), "PFA section inaccessible"
            Set c = rows.Controls(UXFieldControl("ReadinessDate"))
            Check c.Visible And rows!uxReadinessCalendar.Visible, "PFA date or its calendar button hidden"
            c.SetFocus: c.Text = "11/26"
            rows!txtCounselingDate.SetFocus
            Check UXSurfaceEdited("ReadinessDate"), "Typed MM/YY rejected"
            Check f!txtReadinessDate = "2026-11-01" And c.Value = "11/26", "Typed MM/YY not stored as its month: " & f!txtReadinessDate & " / " & c.Value
            Check UXReadinessDateEntry("3/25") = "2025-03-01" And UXReadinessDateEntry("2026-11-17") = "2026-11-01" And UXReadinessDateEntry(#11/30/2026#) = "2026-11-01", "MM/YY entry rules"
            Check UXReadinessDateEntry("13/26") = "13/26", "Invalid month became a date"
            Check UXReadinessCalendarOpen(), "PFA calendar did not open"
            Check Forms!frmUXCounselingDate!lblHeading.Caption = "20. PFA date", "Calendar not labelled for the PFA date"
            Forms!frmUXCounselingDate!txtDate = #11/17/2026#
            Check UXCounselingCalendarUse(), "PFA calendar choice failed"
            Check f!txtReadinessDate = "2026-11-01" And c.Value = "11/26", "Calendar day in November 2026 not shown as 11/26: " & c.Value
            Check UXCounselingCalendarOpen(), "Block 30 calendar did not open after the PFA calendar"
            Check Forms!frmUXCounselingDate!lblHeading.Caption = "Counseling date", "Block 30 calendar kept the PFA label"
            UXCloseScreen "frmUXCounselingDate"
        Else
            Check Not rows!txtReadinessDate.Visible, kind & " shows a PFA date"
        End If
        Check CheckpointEditor(), "Date checkpoint failed"
        Set values = EditorPendingFields()
        Check values("DateReported") = "2026-09-30", "Checkpoint date differs"
        Check UXSaveEditor(), "Calendar dates not saveable: " & f!uxError.Caption
        Check DLookup("DateReported", "Report", "ReportID=" & id) = "2026-09-30", "Selected date reverted"
        Check UXSelectSection(3), "Counseling inaccessible"
        revision = CurrentRevision(id)
        Check UXCounselingCalendarOpen(), "Separate counseling calendar did not open"
        Check Forms!frmUXCounselingDate!txtDate.ShowDatePicker = 1, "Counseling calendar missing date picker"
        Forms!frmUXCounselingDate!txtDate = #9/29/2026#
        Check UXCounselingCalendarUse(), "Calendar selection failed"
        Check f!txtCounselingDate = "2026-09-29" And CurrentRevision(id) = revision, "Calendar selection did not stay pending"
        Check UXSaveEditor(), "Calendar selection failed to save"
        For Each status In Array("NOT REQ", "NOT PERF")
            stage = kind & " typed " & status
            Set c = rows!txtCounselingDate: c.SetFocus: c.Text = CStr(status)
            f!cmdClose.SetFocus
            Check Nz(f!txtCounselingDate, "") = status, "Typed counseling status rejected"
            Check UXSaveEditor(), "Typed counseling status did not save"
        Next
        rows!txtCounselingDate = "NOT PERF": UXSurfaceEdited "CounselingDate"
        Check UXSaveEditor(), "Calendar formatting rejected counseling status"
        Check DLookup("CounselingDate", "Report", "ReportID=" & id) = "NOT PERF", "Status did not save"
        rows!txtCounselingDate.SetFocus: rows!txtCounselingDate.Text = "BAD"
        f!cmdClose.SetFocus
        Check Not UXSaveEditor(), "Invalid counseling text saved"
        Check UXCounselingCalendarOpen(), "Invalid text blocked date correction"
        Forms!frmUXCounselingDate!txtDate = #9/30/2026#
        Check UXCounselingCalendarUse(), "Calendar could not correct invalid text"
        Check UXSaveEditor(), "Corrected calendar date did not save"
        rows!txtCounselingDate = "NOT PERF": UXSurfaceEdited "CounselingDate"
        Check UXCounselingCalendarOpen(), "Status did not allow calendar opening"
        UXCloseScreen "frmUXCounselingDate"
        Check f!txtCounselingDate = "NOT PERF", "Calendar cancellation changed status"
        Check UXSelectSection(0), "Admin inaccessible"
        revision = CurrentRevision(id)
        ' Clearing block 16 before a save puts back the marks NOB replaced.
        Dim promotionBefore As Variant
        f.Controls(UXFieldControl("Trait33")) = "4.0": promotionBefore = f.Controls(UXFieldControl("PromotionRecommendation")).Value
        rows.Controls(UXFieldControl("ObservationStatus")) = True
        Check UXSurfaceEdited("ObservationStatus"), "Whole NOB toggle on failed"
        Check f.Controls(UXFieldControl("Trait33")) = "NOB", "Whole NOB toggle did not stage NOB"
        rows.Controls(UXFieldControl("ObservationStatus")) = False
        Check UXSurfaceEdited("ObservationStatus"), "Whole NOB toggle off failed"
        Check Nz(f.Controls(UXFieldControl("Trait33")), "") = "4.0" And Nz(f.Controls(UXFieldControl("PromotionRecommendation")), "") = Nz(promotionBefore, ""), "Clearing block 16 did not restore the replaced marks"
        Check Nz(f.Controls(UXFieldControl("Trait34")), "") <> "NOB", "Clearing block 16 left an untouched trait at NOB"
        ' A mark changed after NOB is deliberate: clearing block 16 keeps every mark.
        rows.Controls(UXFieldControl("ObservationStatus")) = True
        Check UXSurfaceEdited("ObservationStatus"), "Whole NOB second toggle failed"
        f.Controls(UXFieldControl("Trait34")) = "2.0"
        rows.Controls(UXFieldControl("ObservationStatus")) = False
        Check UXSurfaceEdited("ObservationStatus"), "Whole NOB second clear failed"
        Check f.Controls(UXFieldControl("Trait33")) = "NOB" And f.Controls(UXFieldControl("Trait34")) = "2.0", "Clearing block 16 overwrote a later mark change"
        Check CurrentRevision(id) = revision, "NOB toggling wrote a revision"
        rows.Controls(UXFieldControl("ObservationStatus")) = True
        Check UXSurfaceEdited("ObservationStatus"), "Whole NOB checkbox failed"
        For Each key In TraitKeys(): Check f.Controls(UXFieldControl(CStr(key))) = "NOB", "Whole NOB left a numeric/blank mark": Next
        Check f.Controls(UXFieldControl("PromotionRecommendation")) = "NOB", "Whole NOB recommendation not set"
        Check CurrentRevision(id) = revision, "Whole NOB checkbox wrote a revision"
        If kind = "CHIEFEVAL" Then Check rows!uxChiefGuidance33.Caption = CatalogueTraitDescription("CHIEFEVAL", 33, "NOB"), "Chief guidance did not follow NOB mark"
        Check UXSaveEditor(), "Whole NOB not saveable"
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
    On Error Resume Next
    If Not c Is Nothing Then c.Undo
End Function

