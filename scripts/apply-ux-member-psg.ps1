# Shared group baseline selections save on change; PDF creation uses real identity.
foreach($m in @('modRSDirectoryUI','modUXResourceTools','modUXTraitGrid','modPDFImport','modCommandDirectory','modUXDirectory','modUXHelp')){
 $app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))
}
$app.DoCmd.OpenForm('frmUXGroupRSCA',1);$f=$app.Forms.Item('frmUXGroupRSCA')
try{$app.DeleteControl($f.Name,'cmdSave')}catch{}
$f.Controls('cboBaseline').AfterUpdate='=RSGroupBaselineChanged()'
$f.Controls('chkExcludes').AfterUpdate='=RSSaveGroupCalculation()'
$f.Controls('lblExcludes').Caption='These reports have not yet been included in the selected BOL totals.'
$f.Controls('chkExcludes').ControlTipText='Confirm the historical RSCA and graded trait count exclude every report in this summary group. Swift adds this group once.'
$f.Controls('cmdStage').Caption='Apply reporting senior details to reports'
$f.Controls('cmdStage').Left=400;$f.Controls('cmdStage').Width=12600
$app.DoCmd.Close(2,$f.Name,1)
$app.DoCmd.OpenForm('frmUXDirectory',1);$f=$app.Forms.Item('frmUXDirectory')
$f.Controls('lblScope').Caption='Members link by SSN or a unique matching name without a conflicting SSN. Ambiguous names require confirmation. Profile edits do not rewrite reports.'
$app.DoCmd.Close(2,$f.Name,1)
