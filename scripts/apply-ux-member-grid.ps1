foreach($m in @('modUXFeedback','modUXTraitGrid')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
$app.DoCmd.OpenForm('frmUXTraitRows',1);$f=$app.Forms.Item('frmUXTraitRows')
foreach($c in $f.Controls){if($c.Visible){$c.Left=[int]($c.Left+700);$c.Tag=[string]$c.Left+':'+[string]$c.Width}}
[void](Control $f 104 'uxOpenReport' 30 40 650 360 'Open' '=UXGridOpenReport()');$script:swiftControl.Tag='30:650'
$f.Width=13950
SaveScreen $f 'frmUXTraitRows'
$app.DoCmd.OpenForm('frmUXGroups',1);$f=$app.Forms.Item('frmUXGroups');$f.Controls('cmdOpen').OnClick='=UXOpenGroupMembers()';$f.Controls('lstGroups').OnDblClick='=UXOpenGroupMembers()';SaveScreen $f 'frmUXGroups'
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor');$f.Controls('cmdClose').OnClick='=UXFeedbackCloseEditor()';SaveScreen $f 'frmEditor'
$app.DoCmd.OpenForm('frmUXComparison',1);$f=$app.Forms.Item('frmUXComparison');$f.Controls('lblHeading').Caption='Group members & comparison';$f.Controls('cmdBack').Caption='Group tools';SaveScreen $f 'frmUXComparison'
