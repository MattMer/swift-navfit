# Live diagnostics and source-backed report/group issue list. Schema unchanged.
foreach($module in @('modLiveValidation','modUXValidation','modRecovery','modUX','modUXDashboard','modValidation','modUXSurface','modUXHelp','modSwift')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}
NewScreen 'frmUXValidation' 'Validation issues' 10000;$f=$script:swiftForm
$f.OnLoad='=UXLoadValidation()';$f.OnResize='=UXValidationLayout()';$f.OnTimer='=UXLoadValidation()';$f.TimerInterval=1000
[void](Control $f 100 'lblBasis' 400 950 12600 750 '')
[void](Control $f 100 'lblCount' 400 1750 12600 550 '')
[void](Control $f 104 'cmdPrevious' 10300 1800 1200 450 'Previous' '=UXValidationPage(-1)')
[void](Control $f 104 'cmdNext' 11700 1800 1200 450 'Next' '=UXValidationPage(1)')
[void](Control $f 110 'lstIssues' 400 2400 12600 2700);$c=$script:swiftControl;$c.ColumnCount=5;$c.ColumnHeads=$true;$c.ColumnWidths='0;2400;3000;1900;4900';$c.AfterUpdate='=UXValidationSelection()';$c.OnDblClick='=UXValidationOpenField()'
[void](Control $f 109 'txtEvidence' 400 5250 12600 3100);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2
[void](Control $f 104 'cmdOpen' 400 8850 6000 500 'Open affected report block' '=UXValidationOpenField()')
[void](Control $f 104 'cmdBack' 6800 8850 6000 500 'Back to workspace' '=UXValidationClose()')
$i=0;foreach($name in @('cmdPrevious','cmdNext','lstIssues','txtEvidence','cmdOpen','cmdBack')){$f.Controls($name).TabIndex=$i;$i++}
SaveScreen $f 'frmUXValidation'
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
# Keep the existing responsive action-area identity and geometry; make status keyboard actionable.
$c=$f.Controls('uxDirty');$x=$c.Left;$y=$c.Top;$w=$c.Width;$h=$c.Height
$app.DeleteControl('frmEditor','uxDirty')
[void](Control $f 104 'uxDirty' $x $y $w $h 'Validation issues' '=UXOpenValidation()');$script:swiftControl.TabStop=$true
$f.Controls('cmdChecklist').Caption='Validation issues';$f.Controls('cmdChecklist').OnClick='=UXOpenValidation()'
$f.OnTimer='=UXLiveTimer()';$f.TimerInterval=1000
SaveScreen $f 'frmEditor'
$app.DoCmd.OpenForm('frmUXEditorFields',1);$f=$app.Forms.Item('frmUXEditorFields')
foreach($c in $f.Controls){if($c.Tag -like 'SURF:*' -and $c.ControlType -in @(109,111)){$c.OnChange='=UXLiveTyped()'}}
SaveScreen $f 'frmUXEditorFields'
$app.DoCmd.OpenForm('frmMajorWorkbench',1);$f=$app.Forms.Item('frmMajorWorkbench')
[void](Control $f 104 'uxValidation' 2850 7250 10450 450 'All summary-group validation issues' '=UXOpenValidation(True)');$script:swiftControl.Tag='WBX:Members:Defaults:Review:Comparison:History'
SaveScreen $f 'frmMajorWorkbench'
