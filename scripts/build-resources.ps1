# Trial (forms-as-source): the non-form resources the build generates. Forms
# themselves load from src/forms. The catalogue stages touch no forms; the
# NAVFIT template and report blocks are copied verbatim from apply-ux-navfit
# and apply-pdf-output and would move into shared resource scripts if adopted.
foreach($resourceStage in @('apply-pdf-print-catalogue','apply-ux-field-catalogue','apply-ux-surface-catalogue')){
 Write-Output ('Build resource: '+$resourceStage)
 . (Join-Path $PSScriptRoot ($resourceStage+'.ps1'))
}
Write-Output 'Build resource: modNAVFITTemplate'
$nativeTemplate=Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb'
if((Get-FileHash -LiteralPath $nativeTemplate).Hash -cne 'DD3806E2DAAAE54B3F9B4EC07D59FB0654B46AADECEDCFF54F57DCC75167B2FA'){throw 'Review the changed native blank before embedding it'}
$encoded=[Convert]::ToBase64String([IO.File]::ReadAllBytes($nativeTemplate))
$templateCode=[Collections.Generic.List[string]]::new();$templateCode.Add('Attribute VB_Name = "modNAVFITTemplate"');$templateCode.Add('Option Explicit')
$templateCode.Add('Public Function NAVFITTemplateHash() As String');$templateCode.Add(' NAVFITTemplateHash = "'+(Get-FileHash -LiteralPath $nativeTemplate).Hash+'"');$templateCode.Add('End Function')
$chunks=[int][math]::Ceiling($encoded.Length/16384)
$templateCode.Add('Public Function NAVFITTemplateData() As String');$templateCode.Add(' Dim result As String')
for($i=0;$i -lt $chunks;$i++){$templateCode.Add(' result = result & NativeTemplateChunk'+$i+'()')}
$templateCode.Add(' NAVFITTemplateData = result');$templateCode.Add('End Function')
for($i=0;$i -lt $chunks;$i++){
 $part=$encoded.Substring($i*16384,[math]::Min(16384,$encoded.Length-$i*16384))
 $templateCode.Add('Private Function NativeTemplateChunk'+$i+'() As String');$templateCode.Add(' Dim result As String')
 for($offset=0;$offset -lt $part.Length;$offset+=512){$templateCode.Add(' result = result & "'+$part.Substring($offset,[math]::Min(512,$part.Length-$offset))+'"')}
 $templateCode.Add(' NativeTemplateChunk'+$i+' = result');$templateCode.Add('End Function')
}
$templateModule=Join-Path $root 'build/modNAVFITTemplate.bas';[IO.File]::WriteAllLines($templateModule,$templateCode,[Text.UTF8Encoding]::new($false));$app.LoadFromText(5,'modNAVFITTemplate',$templateModule)
Write-Output 'Build resource: rptPDFOutput'
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
Write-Output 'Build resource: database startup properties'
try{$app.CurrentDb().Properties('StartupShowDBWindow').Value=$false}catch{$db=$app.CurrentDb();$p=$db.CreateProperty('StartupShowDBWindow',1,$false);$db.Properties.Append($p)}
