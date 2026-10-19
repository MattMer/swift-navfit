# Saving, pre-edit scope and contextual-action presentation; no report-field geometry changes.
foreach($m in @('modUX','modSwift','modRecovery','modUXValidation','modUXTraitGrid','modUXRecursiveDefaults','modUXHeaderReview','modUXReview','modUXGroupLifecycle')){
 $app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))
}
$app.DoCmd.OpenForm('frmUXBulk',1);$f=$app.Forms.Item('frmUXBulk')
if(-not @($f.Controls|ForEach-Object {$_.Name}).Contains('uxScope')){
 $scope=$app.CreateControl('frmUXBulk',100,0,'','',2850,2500,10450,400);$scope.Name='uxScope'
}
$scope=$f.Controls.Item('uxScope');$scope.Caption='Scope is checked before editing and again before confirmation.'
$scope.FontName='Segoe UI';$scope.FontSize=9;$scope.ForeColor=2501670;$scope.BackStyle=0
$app.DoCmd.Close(2,'frmUXBulk',1)
$app.DoCmd.OpenForm('frmUXSharedFields',1);$f=$app.Forms.Item('frmUXSharedFields')
foreach($c in $f.Controls){
 if($c.Name -like 'value*' -and $c.ControlType -in @(106,109,111)){$c.OnEnter='=UXSharedScope("'+$c.Name.Substring(5)+'")'}
}
$app.DoCmd.Close(2,'frmUXSharedFields',1)
Write-Output 'Save/recovery messages, pre-edit scope and contextual actions installed.'
