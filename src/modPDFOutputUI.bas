Attribute VB_Name = "modPDFOutputUI"
Option Explicit
Private previewGroup As Long, previewReport As Long
Private Const PDF_OVERFLOW_HELP As String = "Overflow continues on additional pages. An asterisk marks a single-line field whose full value is on an additional page."

Public Function PDFOutputOpenReport() As Boolean
    On Error GoTo Failed
    If PDFOutputBusy() Then Exit Function
    previewGroup = 0: previewReport = Nz(Forms!frmEditor!txtReportID, 0)
    If CurrentProject.AllForms("frmPDFOutput").IsLoaded Then DoCmd.Close acForm, "frmPDFOutput", acSaveNo
    DoCmd.OpenForm "frmPDFOutput": PDFOutputOpenReport = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Create PDFs"
End Function
Public Function PDFOutputOpenGroup(ByVal screen As String) As Boolean
    On Error GoTo Failed
    If PDFOutputBusy() Then Exit Function
    previewGroup = UXResolvedGroupContext(screen): previewReport = 0
    If previewGroup = 0 Then Exit Function
    If CurrentProject.AllForms("frmPDFOutput").IsLoaded Then DoCmd.Close acForm, "frmPDFOutput", acSaveNo
    DoCmd.OpenForm "frmPDFOutput": PDFOutputOpenGroup = True: Exit Function
Failed:
    MsgBox Err.Description, vbExclamation, "Create group PDFs"
End Function
Private Sub PDFOutputLetterChoices()
    PDFOutputSetLetter Nz(Forms!frmPDFOutput!cboISIC, 0), Nz(Forms!frmPDFOutput!cboOnTime, "")
End Sub
Public Function PDFOutputLoad() As Boolean
    With Forms!frmPDFOutput
        !cboSource = "Current drafts": !cboFormat = IIf(previewGroup > 0, "Combined packet", "Individual PDFs")
        !chkSummary = (previewGroup > 0): !chkSummary.Enabled = (previewGroup > 0)
        !cboISIC.RowSource = "SELECT ISICID,Title FROM ISIC WHERE IsArchived=False ORDER BY Title"
        !cboISIC = IIf(GroupISIC(previewGroup) > 0, GroupISIC(previewGroup), Null): !cboOnTime = Null
        !cboISIC.Enabled = (previewGroup > 0): !cboOnTime.Enabled = (previewGroup > 0)
        !txtFolder = PDFOutputFolderPreference(): !cmdOpenFolder.Enabled = False: !cmdOpenPDF.Enabled = False
    End With
    PDFOutputLoad = PDFOutputRefreshPreview()
End Function
Public Function PDFOutputRefreshPreview() As Boolean
    Dim current As Object, row As Object, text As String, i As Long
    On Error GoTo Failed
    PDFOutputCapture previewGroup, previewReport, (Forms!frmPDFOutput!cboSource = "Current drafts")
    Set current = PDFOutputJob()
    For Each row In current("Reports")
        If text <> "" Then text = text & ";"
        text = text & row("ID") & ";" & UXListCell(row("Caption")) & ";" & UXListCell(row("Kind")) & ";" & UXListCell(row("Source")) & ";" & UXListCell(PDFOutputFileName(current("Title"), row))
    Next
    With Forms!frmPDFOutput
        !lstReports.RowSource = text
        !lblScope.Caption = current("Title") & " | " & current("Reports").Count & " reports | Alphabetical by last name"
        !txtFindings = "Captured " & Format$(current("Captured"), "yyyy-mm-dd h:nn:ss AM/PM") & ". Current-draft mode reads pending values without saving or discarding them. Pending group cells take precedence for the same field; both buffers remain unchanged." & vbCrLf & "Findings, missing values and report status never prevent PDF generation. " & PDF_OVERFLOW_HELP & vbCrLf & "Signatures, unrecorded qualifications/ranking/RSCA and submission acceptance remain blank. A computed PSG RSCA is not substituted for an official report RSCA. Choose a protected export folder; PDFs may contain identifiers."
        !lblStatus.Caption = "Preview captured. All listed reports will be included. Choose the output format, then create PDFs."
        !cmdGenerate.Enabled = True: !cmdOpenFolder.Enabled = False: !cmdOpenPDF.Enabled = False
    End With
    PDFOutputRefreshPreview = True: Exit Function
Failed:
    Forms!frmPDFOutput!lblStatus.Caption = "Preview unavailable: " & Err.Description
End Function
Public Function PDFOutputSelectRow() As Boolean
    Dim current As Object, row As Object
    On Error GoTo Failed
    Set current = PDFOutputJob()
    For Each row In current("Reports")
        If row("ID") = Nz(Forms!frmPDFOutput!lstReports, -1) Then
            Forms!frmPDFOutput!txtFindings = row("Caption") & " | " & row("Source") & vbCrLf & "Validation findings are informational; PDFs can be generated with these issues. " & PDF_OVERFLOW_HELP & vbCrLf & row("Findings")
            Exit For
        End If
    Next
    PDFOutputSelectRow = True: Exit Function
Failed:
    Forms!frmPDFOutput!lblStatus.Caption = Err.Description
End Function
Public Function PDFOutputBrowseFolder() As Boolean
    Dim dialog As Object, folder As String
    On Error GoTo Failed
    Set dialog = Application.FileDialog(4)
    dialog.Title = "Choose the folder for this PDF export"
    If dialog.Show <> -1 Then Exit Function
    folder = dialog.SelectedItems(1)
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(folder) Then Err.Raise 76, "modPDFOutputUI.PDFOutputBrowseFolder", "Choose an existing folder."
    Forms!frmPDFOutput!txtFolder = folder: PDFOutputRememberFolder folder
    PDFOutputBrowseFolder = True: Exit Function
Failed:
    Forms!frmPDFOutput!lblStatus.Caption = "Folder selection failed: " & Err.Description
End Function
Public Function PDFOutputPagePreview() As Boolean
    On Error GoTo Failed
    If Nz(Forms!frmPDFOutput!txtFolder, "") = "" Then
        If Not PDFOutputBrowseFolder() Then Exit Function
    End If
    PDFOutputLetterChoices
    PDFOutputPreviewPages CStr(Forms!frmPDFOutput!txtFolder), Nz(Forms!frmPDFOutput!chkSummary, False)
    Forms!frmPDFOutput.Visible = False
    PDFOutputPagePreview = True: Exit Function
Failed:
    Forms!frmPDFOutput!lblStatus.Caption = "Page preview: " & Err.Number & " " & Err.Description
End Function
Public Function PDFOutputCreate() As Boolean
    Dim selected As Object, current As Object, row As Object, folder As String, f As Access.Form, failure As Long, message As String
    On Error GoTo Failed
    Set f = Forms!frmPDFOutput
    If Nz(f!txtFolder, "") = "" Then
        If Not PDFOutputBrowseFolder() Then Exit Function
    End If
    Set selected = CreateObject("Scripting.Dictionary")
    Set current = PDFOutputJob()
    For Each row In current("Reports"): selected(CStr(row("ID"))) = True: Next
    f!txtFindings.SetFocus
    f!cmdGenerate.Enabled = False: f!cmdRefresh.Enabled = False: f!cmdClose.Enabled = False
    f!lblStatus.Caption = "Creating print-ready PDFs from the captured preview...": f.Repaint
    PDFOutputLetterChoices
    folder = PDFOutputGenerate(Nz(f!txtFolder, ""), selected, Nz(f!cboFormat, ""), Nz(f!chkSummary, False))
    'The group remembers its ISIC for the next summary letter.
    If previewGroup > 0 Then SetGroupISIC previewGroup, Nz(f!cboISIC, 0)
    f!lblStatus.Caption = "Completed: " & folder
    f!cmdOpenFolder.Enabled = True: f!cmdOpenPDF.Enabled = True
    f!cmdGenerate.Enabled = True: f!cmdRefresh.Enabled = True: f!cmdClose.Enabled = True
    PDFOutputCreate = True: Exit Function
Failed:
    failure = Err.Number: message = Err.Description
    If Not f Is Nothing Then
        f!lblStatus.Caption = "PDF generation failed: " & failure & " " & message
        f!cmdGenerate.Enabled = True: f!cmdRefresh.Enabled = True: f!cmdClose.Enabled = True
    End If
End Function
Public Function PDFOutputOpenResult(ByVal folderOnly As Boolean) As Boolean
    Dim current As Object, path As String
    On Error GoTo Failed
    Set current = PDFOutputJob(): path = current("Folder")
    If path = "" Then Exit Function
    If Not folderOnly Then path = path & "\" & current("Files")(1)
    CreateObject("Shell.Application").ShellExecute path, "", "", "open", 1
    PDFOutputOpenResult = True: Exit Function
Failed:
    Forms!frmPDFOutput!lblStatus.Caption = "Unable to open output: " & Err.Description
End Function
Public Function PDFOutputClose() As Boolean
    If PDFOutputBusy() Then Exit Function
    PDFOutputEndPreview
    DoCmd.Close acForm, "frmPDFOutput", acSaveNo: PDFOutputClose = True
End Function
