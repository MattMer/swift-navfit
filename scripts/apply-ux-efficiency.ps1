if(-not $script:swiftManagedBuild){. (Join-Path $PSScriptRoot 'ensure-native-unload.ps1')}
# Remove the duplicate, hidden comparison list after adapters use the editable grid.
$app.DoCmd.OpenForm('frmUXComparison',1);$f=$app.Forms.Item('frmUXComparison')
if(@($f.Controls|ForEach-Object {$_.Name}).Contains('lstTraits')){$app.DeleteControl('frmUXComparison','lstTraits')}
$app.DoCmd.Close(2,'frmUXComparison',1)
# Legacy editor inputs also notify the event-driven recovery/validation timer.
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
foreach($c in $f.Controls){if($c.ControlType -in @(109,111) -and ($c.Name -like 'txt*' -or $c.Name -eq 'cboKind' -or $c.Name -like 'cboTrait*' -or $c.Name -eq 'cboObservationStatus' -or $c.Name -eq 'cboPromotionRecommendation')){$c.OnChange='=UXLiveTyped()'}}
$app.DoCmd.Close(2,'frmEditor',1)
