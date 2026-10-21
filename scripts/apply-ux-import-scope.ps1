if(-not $script:swiftManagedBuild){foreach($m in @('modImportReview','modSwift','modChecklist')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
$app.DoCmd.OpenForm('frmUXImportChoose',1);$f=$app.Forms.Item('frmUXImportChoose');$f.Controls('cmdNAVFIT').Caption='Choose NAVFIT database - full report mapping';SaveScreen $f 'frmUXImportChoose'
