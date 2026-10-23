param([switch]$RefreshBackgroundsOnly)
# Manifest-owned Access report and preview. All print backgrounds are embedded.
. (Join-Path $PSScriptRoot 'access-properties.ps1')
function SetPDFBackgrounds($report,[bool]$create) {
 $layout=Get-Content (Join-Path $root 'src/pdf-print-layout.json') -Raw|ConvertFrom-Json
 foreach($template in $layout.templates){foreach($page in $template.pages){
  $name='bg_'+$template.kind+'_'+$page.page
  if($create){$c=$app.CreateReportControl($report.Name,103,0,'','',0,0,12240,15840);$c.Name=$name}else{$c=$report.Controls($name)}
  $c.SizeMode=1;$c.PictureType=0;$c.BorderStyle=0
  $picturePath=Join-Path $root ('src/pdf-templates/'+$page.image)
  $c.Picture=[string]$picturePath;$c.Visible=$false
 }}
}
if($RefreshBackgroundsOnly){
 $app.DoCmd.OpenReport('rptPDFOutput',1);$r=$app.Reports.Item('rptPDFOutput')
 SetPDFBackgrounds $r $false
 $app.DoCmd.Close(3,'rptPDFOutput',1)
 return
}
function AddPDFControl($form,$type,$name,$x,$y,$w,$h,$caption='',$event='') {
 $c=$app.CreateControl($form.Name,$type,0,'','',$x,$y,$w,$h);$c.Name=$name
 if($type -in @(100,104)){$c.Caption=$caption}
 if($type -ne 106){$c.FontName='Segoe UI';$c.FontSize=10}
 if($event){$c.OnClick=$event}
 $script:pdfControl=$c
}
$app.LoadFromText(2,'frmPDFOutput',(Join-Path $root 'src/BlankForm.txt'))
$app.DoCmd.OpenForm('frmPDFOutput',1);$f=$app.Forms.Item('frmPDFOutput')
$f.Caption='Preview and create PDFs';$f.Width=14200;$f.Section(0).Height=10300
$f.PopUp=$true;$f.Modal=$false;$f.AutoCenter=$true;$f.RecordSelectors=$false;$f.NavigationButtons=$false;$f.ScrollBars=2
AddPDFControl $f 100 'lblScope' 300 250 13500 500 'PDF preview'
AddPDFControl $f 100 'lblSource' 300 900 1800 300 'Capture values'
AddPDFControl $f 111 'cboSource' 2100 850 3400 400
$pdfControl.RowSourceType='Value List';$pdfControl.RowSource='Current drafts;Saved revisions';$pdfControl.LimitToList=$true;$pdfControl.AfterUpdate='=PDFOutputRefreshPreview()'
AddPDFControl $f 100 'lblFormat' 5800 900 1400 300 'Output'
AddPDFControl $f 111 'cboFormat' 7200 850 3400 400
$pdfControl.RowSourceType='Value List';$pdfControl.RowSource='Combined packet;Individual PDFs;Both';$pdfControl.LimitToList=$true
AddPDFControl $f 106 'chkSummary' 11000 880 300 300
AddPDFControl $f 100 'lblSummary' 11400 900 2300 400 'Summary letter first'
AddPDFControl $f 100 'lblFolder' 300 1450 1800 300 'Export folder'
AddPDFControl $f 109 'txtFolder' 2100 1400 9000 450;$pdfControl.Locked=$true
AddPDFControl $f 104 'cmdBrowse' 11400 1400 2400 450 'Choose folder' '=PDFOutputBrowseFolder()'
AddPDFControl $f 100 'lblReports' 300 2100 13500 350 'Report order: last name, first name, middle initial. Select a row to view informational findings.'
AddPDFControl $f 110 'lstReports' 300 2550 13500 2700
$pdfControl.RowSourceType='Value List';$pdfControl.ColumnCount=5;$pdfControl.ColumnWidths='0;2800;1100;2900;6500';$pdfControl.AfterUpdate='=PDFOutputSelectRow()'
AddPDFControl $f 109 'txtFindings' 300 5450 13500 2200;$pdfControl.Locked=$true;$pdfControl.ScrollBars=2
AddPDFControl $f 100 'lblStatus' 300 7850 13500 800 'Capture the preview before creating PDFs.'
AddPDFControl $f 104 'cmdPagePreview' 300 9500 6100 500 'Preview printed pages (close preview to return)' '=PDFOutputPagePreview()'
AddPDFControl $f 104 'cmdRefresh' 300 8800 2700 500 'Refresh captured preview' '=PDFOutputRefreshPreview()'
AddPDFControl $f 104 'cmdGenerate' 3200 8800 3200 500 'Create print-ready PDFs' '=PDFOutputCreate()'
AddPDFControl $f 104 'cmdOpenPDF' 6600 8800 2300 500 'Open first PDF' '=PDFOutputOpenResult(False)'
AddPDFControl $f 104 'cmdOpenFolder' 9100 8800 2300 500 'Open output folder' '=PDFOutputOpenResult(True)'
AddPDFControl $f 104 'cmdClose' 11600 8800 2200 500 'Close' '=PDFOutputClose()'
$f.OnLoad='=PDFOutputLoad()';$f.OnUnload='[Event Procedure]'
$f.Module.AddFromString("Private Sub Form_Unload(Cancel As Integer)`r`n Cancel = PDFOutputBusy()`r`n If Not Cancel Then PDFOutputEndPreview`r`nEnd Sub")
$app.DoCmd.Close(2,'frmPDFOutput',1)
Write-Output 'PDF preview form installed'
try{$app.DoCmd.DeleteObject(3,'rptPDFOutput')}catch{}
$r=$app.CreateReport();$old=$r.Name;$r.Width=12240;$r.Section(0).Height=15840
$r.Section(0).ForceNewPage=1;$r.Section(0).KeepTogether=$false
foreach($s in @(1,2,3,4)){try{$r.Section($s).Height=0}catch{}}
$r.Printer.PaperSize=1;$r.Printer.Orientation=1
$r.Printer.LeftMargin=0;$r.Printer.RightMargin=0;$r.Printer.TopMargin=0;$r.Printer.BottomMargin=0
SetPDFBackgrounds $r $true
Write-Output 'PDF backgrounds embedded'
$pdfSlotMatch=[regex]::Match((Get-Content (Join-Path $root 'src/modPDFOutput.bas') -Raw),'(?m)^Public Const PDF_OUTPUT_SLOTS As Long = (\d+)')
if(-not $pdfSlotMatch.Success){throw 'Missing PDF text-slot contract'}
$pdfSlotCount=[int]$pdfSlotMatch.Groups[1].Value
if($pdfSlotCount -lt 1 -or $pdfSlotCount -gt 500){throw 'PDF text-slot count exceeds the supported native report design'}
for($i=1;$i -le $pdfSlotCount;$i++){
 $suffix=$i.ToString('000');$c=$app.CreateReportControl($old,109,0,'','',0,0,200,200)
 $c.Name='v'+$suffix;$c.BackStyle=0;$c.BorderStyle=0;$c.CanGrow=$false;$c.CanShrink=$false
 $c.TopMargin=0;$c.BottomMargin=0;$c.LeftMargin=0;$c.RightMargin=0;$c.Visible=$false
}
Write-Output 'PDF text slots installed'
$c=$app.CreateReportControl($old,109,0,'','PageNo',0,0,200,200);$c.Name='txtPageNo';$c.Visible=$false
$r.OnOpen='[Event Procedure]';$r.OnClose='=PDFOutputReportClosed()';$r.Section(0).OnFormat='[Event Procedure]'
$r.Module.AddFromString("Private Sub Report_Open(Cancel As Integer)`r`n Cancel = Not PDFOutputReportOpen(Me)`r`nEnd Sub`r`nPrivate Sub Detail_Format(Cancel As Integer, FormatCount As Integer)`r`n Cancel = Not PDFOutputFormat(Me)`r`nEnd Sub")
$app.DoCmd.Close(3,$old,1);$app.DoCmd.Rename('rptPDFOutput',3,$old)
Write-Output 'PDF report installed'
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
try{$app.DeleteControl($f.Name,'uxCreatePDF')}catch{}
AddPDFControl $f 104 'uxCreatePDF' 13400 450 1200 400 'Create PDF' '=PDFOutputOpenReport()'
$app.DoCmd.Close(2,$f.Name,1)
foreach($screen in @('frmUXComparison','frmUXBulk','frmMajorWorkbench')){
 $app.DoCmd.OpenForm($screen,1);$f=$app.Forms.Item($screen)
 try{$app.DeleteControl($screen,'uxCurrentPDF')}catch{}
 $f.Controls('uxCurrentSenior').Width=3300;$f.Controls('uxCurrentExport').Left=6350;$f.Controls('uxCurrentExport').Width=3300
 AddPDFControl $f 104 'uxCurrentPDF' 9850 $f.Controls('uxCurrentExport').Top 3450 430 'Preview / create PDFs' ('=PDFOutputOpenGroup("'+$screen+'")')
 $pdfControl.Tag='GroupFooter';$app.DoCmd.Close(2,$screen,1)
}
