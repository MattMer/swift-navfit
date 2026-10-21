# Shared validation services and contextual access to the complete export findings.
foreach($m in @('modValidationExpansion','modValidationReferences','modSavedReadiness','modLiveValidation','modValidation','modUXValidation','modNAVFITExchange','modUXNAVFIT','modUXHelp')){
if(-not $script:swiftManagedBuild){ $app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
}
$app.DoCmd.OpenForm('frmNAVFITExchange',1);$f=$app.Forms.Item('frmNAVFITExchange')
$bottom=10000
foreach($c in $f.Controls){if($c.Visible -and $c.Name -notin @('cmdReportReadiness','cmdGroupReadiness')){$bottom=[math]::Max($bottom,$c.Top+$c.Height)}}
foreach($name in @('cmdReportReadiness','cmdGroupReadiness')){try{$app.DeleteControl($f.Name,$name)}catch{}}
for($i=0;$i -lt 2;$i++){
 $c=$app.CreateControl($f.Name,104,0,'','',(400+$i*6500),($bottom+200),6000,500)
 $c.Name=@('cmdReportReadiness','cmdGroupReadiness')[$i];$c.Caption=@('All report readiness issues','All group readiness issues')[$i]
 $c.OnClick=@('=UXExportReadiness(False)','=UXExportReadiness(True)')[$i];$c.FontName='Segoe UI';$c.FontSize=10
 $c.BackColor=$f.Controls('cmdExportGroup').BackColor;$c.ForeColor=$f.Controls('cmdExportGroup').ForeColor
}
$f.Section(0).Height=$bottom+1000;$app.DoCmd.Close(2,$f.Name,1)
