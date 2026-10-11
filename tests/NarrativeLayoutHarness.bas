Attribute VB_Name = "NarrativeLayoutHarness"
Option Explicit
Private checks As Long
Private Sub Check(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 8061, , message
    checks = checks + 1
End Sub
Public Function TestNarrativeLayout() As String
    On Error GoTo Failed
    Dim kind As Variant, id As Long, f As Form, surface As Form, c As Control
    Dim width As Variant, value As String, revision As Long
    ConfigureWorkspace "SYNTHETIC NARRATIVE LAYOUT", "Tester"
    value = "Mixed CASE; quotes ""test"" and apostrophe '" & vbCrLf & String$(6000, "x")
    For Each kind In Array("EVAL", "CHIEFEVAL", "FITREP")
        id = CommitDraft(0, CStr(kind), "Synthetic", "Synthetic", Null, Null, "", "Layout test", value)
        revision = DLookup("CurrentRevisionID", "Report", "ReportID=" & id)
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Set f = Forms!frmEditor: f.TimerInterval = 0
        Check UXSelectSection(5), "Comments section failed"
        Set surface = f!uxFieldSurface.Form: Set c = surface!txtOfficialComments
        For Each width In Array(28000, 20000, 14000)
            Check UXResizeEditor(CLng(width), 11000), "Resize failed"
            Check c.Width = UXNarrativeWidth(CStr(kind)), "Narrative width changed with viewport"
            Check c.Height = UXNarrativeHeight(CStr(kind)), "Narrative height changed with viewport"
            Check c.FontName = "Courier New", "Font not locked"
            Check c.FontSize = UXNarrativeFontSize(CStr(kind)), "Point size differs"
            Check c.TextFormat = 0 And Not c.FontBold And Not c.FontItalic, "Rich styling remained"
            Check c.LeftMargin = 90 And c.RightMargin = 90, "Margins differ"
            Check surface.Controls("errOfficialComments").Top >= c.Top + c.Height, "Error overlaps narrative"
            Check c.Left + c.Width <= f!uxFieldSurface.Width - 300, "Narrative exceeds surface"
            Check c.Value = value, "Layout truncated overflow"
        Next
        Check UXResizeEditor(9000, 8000), "Narrow layout failed"
        Check c.Width < UXNarrativeWidth(CStr(kind)), "Narrow viewport did not fit"
        Check InStr(surface!uxNarrativeCount.Caption, "Narrow view") > 0, "Narrow wrapping not disclosed"
        Check c.Value = value, "Narrow layout lost text"
        Check UXSelectSection(0), "Admin section failed"
        Check UXSelectSection(5), "Comments return failed"
        Check c.Value = value, "Section change lost text"
        Check DLookup("CurrentRevisionID", "Report", "ReportID=" & id) = revision, "Layout created revision"
        c = value & vbCrLf & "Pending tail"
        Check UXSurfaceEdited("OfficialComments"), "Visible edit failed"
        Check UXSaveEditor(), "Narrative save failed"
        Check DLookup("OfficialComments", "Report", "ReportID=" & id) = value & vbCrLf & "Pending tail", "Save truncated overflow"
        f.OnUnload = "": DoCmd.Close acForm, "frmEditor", acSaveNo
        DoCmd.OpenForm "frmEditor", , , , , , CStr(id)
        Check Forms!frmEditor!txtOfficialComments = value & vbCrLf & "Pending tail", "Reopen lost text"
        Forms!frmEditor.OnUnload = "": DoCmd.Close acForm, "frmEditor", acSaveNo
    Next
    TestNarrativeLayout = "PASS: " & checks & " narrative font, geometry, viewport, overflow and save/reopen assertions"
    Exit Function
Failed:
    TestNarrativeLayout = "FAIL: " & kind & ": " & Err.Number & " " & Err.Description
End Function
