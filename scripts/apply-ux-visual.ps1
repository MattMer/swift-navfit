# Semantic palette; retain native event expressions and all workflow authorities.
. (Join-Path $PSScriptRoot 'access-properties.ps1')
if(-not $script:swiftManagedBuild){foreach($m in @('modUXVisual','modUXHeaderReview','modUXReview','modUXHelp','modUXViewport','modUXValidation','modLibraryViews','modUXGroupLifecycle','modUXTraitGrid','modReportPages')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
$names=@($app.CurrentProject.AllForms|ForEach-Object {$_.Name})
foreach($name in $names){
 $app.DoCmd.OpenForm($name,1);$f=$app.Forms.Item($name)
 if(-not $app.Eval('UXApplyVisual(Forms("'+$name+'"))')){throw ('Visual palette failed: '+$name)}
 [void]$app.Eval('UXVisualRailBack("'+$name+'")')
 $old=[string]$f.OnActivate
 if($old -eq '[Event Procedure]'){throw ('Native activation needs explicit preservation: '+$name)}
 if($old -notlike '=UXVisualEvent*'){$old=$old.TrimStart('=').Replace('"','""');$f.OnActivate='=UXVisualEvent("'+$name+'","'+$old+'")'}
 if($name -eq 'frmUXResources'){SetAccessProperty ($f.Controls.Item('cmdSeniors')) 'Top' 4000;SetAccessProperty ($f.Controls.Item('uxDirectory')) 'Top' 4850}
 $app.DoCmd.Close(2,$name,1)
}
Write-Output ('Semantic palette installed: '+$names.Count+' forms.')
