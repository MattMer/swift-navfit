foreach($name in @('modUXRecursiveDefaults','modUXBlockTools','modUX','modRSDirectoryUI','modUXHelp','modUXGlobalDefaults')){$app.LoadFromText(5,$name,(Join-Path $root ('src/'+$name+'.bas')))}
$app.DoCmd.OpenForm('frmUXSharedFields',1);$f=$app.Forms.Item('frmUXSharedFields')
foreach($name in @($f.Controls|Where-Object {$_.Name -match '^(state|reset|clear)'}|ForEach-Object {$_.Name})){$app.DeleteControl($f.Name,$name)}
foreach($c in $f.Controls){if($c.Name -like 'value*' -and $c.ControlType -in @(109,111)){$c.OnChange=''}}
$app.DoCmd.Close(2,'frmUXSharedFields',1)
$app.DoCmd.OpenForm('frmUXBulk',1);$f=$app.Forms.Item('frmUXBulk')
foreach($name in @('cmdPreview','cmdApply','lstPreview','txtFullValues','txtGroupImpact','lblReason','txtReason')){if(@($f.Controls|ForEach-Object {$_.Name}).Contains($name)){$app.DeleteControl($f.Name,$name)}}
foreach($name in @($f.Controls|Where-Object {($_.ControlType -eq 100 -and $_.Tag -like 'BLX:*') -or $_.Name -match '^uxSection\dPage'}|ForEach-Object {$_.Name})){$app.DeleteControl($f.Name,$name)}
foreach($c in $f.Controls){if($c.Name -like 'uxValue*' -and $c.ControlType -in @(109,111)){$c.OnChange=''}}
$f.Caption='Defaults and shared fields';$f.Controls('uxStatus').Caption='Edit a field to confirm and apply it.'
$f.Controls('uxStatus').Top=7500;$f.Controls('uxStatus').Height=650
$f.Controls('cmdBack').Top=8250;$f.Controls('cmdBack').Caption='Close'
$f.Section(0).Height=9000
$app.DoCmd.Close(2,'frmUXBulk',1)
foreach($screen in @('frmUXBulk','frmUXComparison','frmMajorWorkbench')){
 $app.DoCmd.OpenForm($screen,1);$f=$app.Forms.Item($screen)
 if(@($f.Controls|ForEach-Object {$_.Name}).Contains('uxGroupTab1')){$f.Controls('uxGroupTab1').Caption='Defaults'}
 $app.DoCmd.Close(2,$screen,1)
}

$app.DoCmd.OpenForm("frmUXGroupRSCA",1)
$app.Forms.Item("frmUXGroupRSCA").Controls("cmdStage").Caption="Apply reporting senior details"
$app.DoCmd.Close(2,"frmUXGroupRSCA",1)
