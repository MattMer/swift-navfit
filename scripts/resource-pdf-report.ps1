param([switch]$RefreshBackgroundsOnly)
# The PDF output report, generated from the tracked page images. It stays
# generated rather than stored as text: its printer blob differs on every save.
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
