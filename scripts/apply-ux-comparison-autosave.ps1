. (Join-Path $PSScriptRoot 'access-properties.ps1')
# Finish bound-control events before the report transaction runs.
$app.LoadFromText(5,'modUXTraitGrid',(Join-Path $root 'src/modUXTraitGrid.bas'))
$app.DoCmd.OpenForm('frmUXTraitRows',1);$f=$app.Forms.Item('frmUXTraitRows')
foreach($key in @('Trait33','Trait34','Trait35','Trait36','Trait37','Trait38','Trait39','PromotionRecommendation','AssignmentRecommendationA','AssignmentRecommendationB')){
    $f.Controls.Item($key).AfterUpdate='=UXGridQueue("'+$key+'")'
}
$app.DoCmd.Close(2,'frmUXTraitRows',1)
$app.DoCmd.OpenForm('frmUXComparison',1);$f=$app.Forms.Item('frmUXComparison')
$f.OnTimer='=UXGridDrain()';$f.TimerInterval=0
if(-not @($f.Controls|ForEach-Object {$_.Name}).Contains('cmdRetryCells')){
 $actionTop=[int]($f.Controls('cmdBack').Top+600)
 [void](Control $f 104 'cmdRetryCells' 2850 $actionTop 4300 450 'Retry pending cells' '=UXGridDrain()');$script:swiftControl.Visible=$false;$script:swiftControl.FontName='Segoe UI';$script:swiftControl.ControlTipText='Retry only unsaved cells; successful earlier revisions are retained.'
 [void](Control $f 104 'cmdDiscardCells' 7500 $actionTop 5500 450 'Discard pending cells' '=UXGridDiscardPending()');$script:swiftControl.Visible=$false;$script:swiftControl.FontName='Segoe UI';$script:swiftControl.ControlTipText='Discard all remaining unsaved cells; previously saved revisions remain.'
 SetAccessProperty ($f.Section(0)) 'Height' ([Math]::Max([int]$f.Section(0).Height,[int]($actionTop+600)))
}
$f.HasModule=$true;$f.OnUnload='[Event Procedure]';$module=$f.Module
$guard='    If Not UXGridUnload(Me) Then Cancel = True'
for($line=1;$line -le $module.CountOfLines;$line++){if($module.Lines($line,1) -match '^\s*UXGridUnload Me\s*$'){$module.ReplaceLine($line,$guard)}}
if($module.Lines(1,$module.CountOfLines) -notmatch 'If Not UXGridUnload\(Me\) Then Cancel = True'){try{$line=$module.ProcStartLine('Form_Unload',0)}catch{$line=$module.CreateEventProc('Unload','Form')};$module.InsertLines($line+1,$guard)}
$app.DoCmd.Close(2,'frmUXComparison',1)
Write-Output 'Comparison autosave controls and unload bindings applied.'
