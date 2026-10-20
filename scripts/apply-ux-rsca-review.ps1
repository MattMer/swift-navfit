# Completion accounting and distinct local report tabs; no schema or field geometry changes.
foreach($m in @('modRSCompletion','modRSDirectory','modGroupPolicy','modUXResourceTools','modReportPages','modUXSeniorTable','modUXHelp')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
$app.DoCmd.OpenForm('frmUXRSCATable',1)
$f=$app.Forms.Item('frmUXRSCATable')
$f.Controls.Item('headdate').Caption='As of (BOL / local completion)'
$app.DoCmd.Close(2,'frmUXRSCATable',1)

