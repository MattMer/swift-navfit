if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modUXImportFlow',(Join-Path $root 'src/modUXImportFlow.bas'))}
NewScreen 'frmUXImportChoose' 'Import reports' 5100;$f=$script:swiftForm;$f.ScrollBars=2
[void](Control $f 100 'lblScope' 400 1200 12600 750 'Choose > Preview > Apply > Results. NAVFIT narrative import creates partial drafts. Structured PDF import can create a draft or merge reviewed values into an existing draft. Both retain source provenance.')
[void](Control $f 104 'cmdNAVFIT' 400 2250 12600 600 'Choose NAVFIT database - narrative-only import' '=ChooseReviewFile()')
[void](Control $f 104 'cmdPDF' 400 3100 12600 600 'Choose structured PDF - supported field mappings' '=OpenScreen("frmPDFImport")')
[void](Control $f 104 'cmdBack' 400 4200 3500 500 'Back to Reports' '=UXNavigate("frmLibrary")')
SaveScreen $f 'frmUXImportChoose'
$app.DoCmd.OpenForm('frmLibrary',1);$f=$app.Forms.Item('frmLibrary')
$f.Controls('uxImport').Caption='Import reports';$f.Controls('uxImport').OnClick='=UXOpenImportChooser()'
SaveScreen $f 'frmLibrary'
NewScreen 'frmPDFImport' 'Import structured PDF' 11200;$f=$script:swiftForm;$f.ScrollBars=2
[void](Control $f 100 'lblInfo' 400 1000 12600 650 'Supplied May 2025 NAVPERS templates; local extraction; no Acrobat required. Expanded mapping remains partial. Blank source fields never clear existing values. Flattened PDFs need manual entry.')
[void](Control $f 104 'cmdChoose' 400 1850 12600 500 'Choose PDF' '=ChoosePDFPreview()')
[void](Control $f 100 'lblCentury' 400 2500 5500 350 'Explicit century for six-digit report dates')
[void](Control $f 111 'cboCentury' 6200 2450 6600 400);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='19;20;21';$c.LimitToList=$true;$c.AfterUpdate='=UXPDFMappingChanged()'
[void](Control $f 100 'lblIdentifier' 400 3000 5500 350 'Combined identifier interpretation')
[void](Control $f 111 'cboIdentifier' 6200 2950 6600 400);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='SSN;Retain only';$c.DefaultValue='"SSN"';$c.LimitToList=$true;$c.AfterUpdate='=UXPDFMappingChanged()'
[void](Control $f 106 'chkLegacy' 400 3500 300 300);$script:swiftControl.DefaultValue='False';$script:swiftControl.AfterUpdate='=UXPDFLegacyChanged()'
[void](Control $f 100 'lblLegacy' 850 3500 12100 400 'Legacy compatibility: allow DoDID mapping for older source files')
[void](Control $f 100 'lblTarget' 400 4050 4000 350 'Merge into draft (leave blank to create new)')
[void](Control $f 111 'cboTarget' 4600 4000 8400 450);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;8000';$c.BoundColumn=1;$c.LimitToList=$true;$c.AfterUpdate='=UXPDFMappingChanged()'
[void](Control $f 100 'lblMergeMode' 400 4550 4000 350 'Existing-draft merge mode')
[void](Control $f 111 'cboMergeMode' 4600 4500 8400 450);$c=$script:swiftControl;$c.RowSourceType='Value List';$c.RowSource='Fill empty fields;Replace mapped values';$c.DefaultValue='"Fill empty fields"';$c.LimitToList=$true;$c.AfterUpdate='=UXPDFMappingChanged()'
[void](Control $f 104 'cmdPreview' 400 5150 6000 450 'Preview mapping and merge' '=UXPreviewPDFImport()')
[void](Control $f 104 'cmdDependency' 6900 5150 6100 450 'Check PDF extraction on this computer' '=UXPDFDependencyCheck()')
[void](Control $f 109 'txtPreview' 400 5800 12600 2900);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2
[void](Control $f 100 'uxState' 400 8850 12600 600 '')
[void](Control $f 104 'cmdImport' 400 9500 6000 500 'Apply reviewed import' '=ImportPDFUI()');$script:swiftControl.Enabled=$false
[void](Control $f 104 'cmdOpenResult' 6900 9500 6100 500 'Open imported report' '=UXOpenImportedReport("PDF")');$script:swiftControl.Enabled=$false
[void](Control $f 100 'lblStatus' 400 10150 12600 650 '')
[void](Control $f 104 'cmdHome' 10400 10850 2600 350 'Back to Reports' '=UXNavigate("frmLibrary")')
$f.OnLoad='=UXPDFStart()';$f.OnUnload='=PDFClearPreview(False)'
[void](Control $f 104 'cmdResume' 400 10850 4600 350 'Resume saved preview' '=UXResumePDFImport()')
[void](Control $f 104 'cmdCancelPreview' 5200 10850 4600 350 'Cancel saved preview' '=UXCancelPDFImport()')
$i=0;foreach($name in @('cmdChoose','cboCentury','cboIdentifier','chkLegacy','cboTarget','cboMergeMode','cmdPreview','cmdDependency','txtPreview','cmdImport','cmdOpenResult','cmdResume','cmdCancelPreview','cmdHome')){$f.Controls($name).TabIndex=$i;$i++}
SaveScreen $f 'frmPDFImport'
$app.DoCmd.OpenForm('frmImportReview',1);$f=$app.Forms.Item('frmImportReview');$f.ScrollBars=2
$f.Controls('lstImportReports').ColumnCount=4;$f.Controls('lstImportReports').ColumnWidths='0;1100;1500;2000'
$f.Controls('cmdImport').Caption='Apply reviewed import';$f.Controls('cmdImport').Enabled=$false
$f.Controls('txtImportMember').AfterUpdate='=UXNarrativeInputChanged()'
$f.Controls('lblScope').Caption='Choose > Preview > Apply > Results. New drafts map the form type and official narrative only; inspect the full mapping preview before Apply. Source field inspection remains available above. Existing-draft reconciliation remains a separate comparison action.'
[void](Control $f 104 'cmdPreview' 400 10800 4200 550 'Preview narrative mapping' '=UXPreviewNarrativeImport()')
$f.Controls('lblScope').Top=900;$f.Controls('lblScope').Height=550
$f.Controls('txtSource').Left=400;$f.Controls('txtSource').Top=2150;$f.Controls('txtSource').Width=4700
$f.Controls('lstImportReports').Top=2900;$f.Controls('lstImportReports').Height=2300
$f.Controls('lstImportFields').Top=2900;$f.Controls('lstImportFields').Height=1900;$f.Controls('lstImportFields').Visible=$false
$f.Controls('lblFieldColumns').Top=2550;$f.Controls('lblFieldColumns').Visible=$false
$f.Controls('txtFieldValue').Left=5500;$f.Controls('txtFieldValue').Top=5050;$f.Controls('txtFieldValue').Width=7500;$f.Controls('txtFieldValue').Height=1350;$f.Controls('txtFieldValue').Visible=$false
$f.Controls('lblImportMember').Left=400;$f.Controls('lblImportMember').Top=5500;$f.Controls('lblImportMember').Width=4700
$f.Controls('txtImportMember').Left=400;$f.Controls('txtImportMember').Top=5900;$f.Controls('txtImportMember').Width=4700
$f.Controls('cmdPreview').Left=5500;$f.Controls('cmdPreview').Top=2150;$f.Controls('cmdPreview').Width=7500
[void](Control $f 104 'cmdInspect' 400 6500 4700 450 'Inspect source fields' '=UXNarrativeInspect()')
$f.Controls('cmdImport').Top=6900;$f.Controls('cmdImport').Left=5500;$f.Controls('cmdImport').Width=3500
$f.Controls('lblImportResult').Top=7550;$f.Controls('lblImportResult').Left=5500;$f.Controls('lblImportResult').Width=7500;$f.Controls('lblImportResult').Height=550
[void](Control $f 109 'uxPreview' 5500 2900 7500 3500);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2
[void](Control $f 100 'uxState' 400 1500 12600 550 '2. Preview - select a source member and preview mapping.')
[void](Control $f 104 'cmdOpenResult' 9400 6900 3600 550 'Open imported report' '=UXOpenImportedReport("NAVFIT")');$script:swiftControl.Enabled=$false
[void](Control $f 100 'lblReconcile' 400 8100 12600 350 'Existing report reconciliation: compare separately before accepting an incoming narrative')
$f.Controls('cboDestinationReport').Top=8550;$f.Controls('cmdCompareIncoming').Top=8550
$f.Controls('cmdLeave').Top=9250;$f.Controls('cmdLeave').Caption='Back to Reports';$f.Controls('cmdLeave').OnClick='=UXLeaveImportReview()'
$f.Section(0).Height=9900;$f.OnUnload='=CloseReviewFile()'
$i=0;foreach($name in @('lstImportReports','txtImportMember','cmdPreview','uxPreview','cmdImport','cmdOpenResult','cmdInspect','lstImportFields','txtFieldValue','cboDestinationReport','cmdCompareIncoming','cmdLeave')){$f.Controls($name).TabIndex=$i;$i++}
SaveScreen $f 'frmImportReview'
