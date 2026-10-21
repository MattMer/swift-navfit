. (Join-Path $PSScriptRoot 'access-properties.ps1')
if(-not $script:swiftManagedBuild){foreach($m in @('modWorkspaceImport','modMigration','modUXNativeGroups')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
$app.DoCmd.OpenForm('frmUXWorkspaceImport',1);$f=$app.Forms.Item('frmUXWorkspaceImport')
SetAccessProperty ($f.Controls.Item('txtPreview')) 'Height' 3000
$c=$app.CreateControl($f.Name,100,0,'','',350,4350,12200,350);$c.Name='lblProgress';$c.Caption='Import activity - progress appears here after you start';$c.FontName='Segoe UI';$c.FontSize=10
$c=$app.CreateControl($f.Name,109,0,'','',350,4750,12200,2400);$c.Name='txtProgress';$c.Locked=$true;$c.ScrollBars=2;$c.FontName='Segoe UI';$c.FontSize=10;$c.ControlTipText='Live copy and verification log. Keep Access open until the result appears.'
$f.HasModule=$true
$code=$f.Module;$code.AddFromString("Private Sub Form_Unload(Cancel As Integer)`r`n Cancel = UXWorkspaceImportBusy()`r`nEnd Sub")
$f.OnUnload='[Event Procedure]'
$app.DoCmd.Close(2,$f.Name,1)
$app.DoCmd.OpenForm('frmUXNativeGroups',1);$f=$app.Forms.Item('frmUXNativeGroups')
SetAccessProperty ($f.Controls.Item('lstGroups')) 'Height' 1450
foreach($v in @(@('cmdSelectAll','Select all for import','=UXSelectAllNativeGroups(True)',400,4000),@('cmdClearSelection','Clear selection','=UXSelectAllNativeGroups(False)',4600,4000))){
 $c=$app.CreateControl($f.Name,104,0,'','',$v[3],3700,$v[4],400);$c.Name=$v[0];$c.Caption=$v[1];$c.OnClick=$v[2];$c.FontName='Segoe UI';$c.FontSize=10;$c.BackColor=$f.Controls.Item('cmdPreview').BackColor;$c.ForeColor=$f.Controls.Item('cmdPreview').ForeColor
}
$app.DoCmd.Close(2,$f.Name,1)
