if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modUXCalculationContext',(Join-Path $root 'src/modUXCalculationContext.bas'))}
$app.DoCmd.OpenForm('frmChecklist',1);$f=$app.Forms.Item('frmChecklist')
[void](Control $f 104 'cmdOpenSection' 8200 8850 4500 550 'Open report section' '=ChecklistOpenSection()');$script:swiftControl.Enabled=$false
$f.Controls('lstFindings').OnDblClick='=ChecklistOpenSection()'
SaveScreen $f 'frmChecklist'
$app.DoCmd.OpenForm('frmUXGroupRSCA',1);$f=$app.Forms.Item('frmUXGroupRSCA')
[void](Control $f 109 'txtCalculationContext' 2866 6750 9954 1500);$c=$script:swiftControl;$c.Locked=$true;$c.ScrollBars=2;$c.ControlTipText='Saved calculation counts, exclusions and manual baseline source'
$f.Controls('lblStatus').Top=8400;$f.Controls('lblStatus').Height=1000
$f.Controls('cmdDirectory').Top=9600;$f.Controls('cmdBack').Top=9600;$f.Section(0).Height=10300;$f.ScrollBars=2
SaveScreen $f 'frmUXGroupRSCA'
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
$f.Controls('uxITA').OnDblClick='=UXEditorCalculationDetails()';$f.Controls('uxITA').ControlTipText='Double-click for pending and saved calculation counts and exclusions'
$f.Controls('uxSGA').OnDblClick='=UXEditorCalculationDetails()';$f.Controls('uxSGA').ControlTipText='Double-click for group calculation counts and exclusions'
$f.Controls('cmdChecklist').Caption='Draft checklist';$f.Controls('cmdChecklist').OnClick='=ChecklistFromEditor()'
[void](Control $f 104 'cmdCalculations' 10000 3350 2400 420 'Calculation details' '=UXEditorCalculationDetails()');$script:swiftControl.Tag='LOCAL:8'
SaveScreen $f 'frmEditor'
