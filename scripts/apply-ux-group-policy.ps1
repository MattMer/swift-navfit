# Group policy context and consequence previews, using existing schema-21 storage.
foreach($m in @('modGroupPolicy','modLiveValidation','modSummaryGroups','modUX','modUXBlockTools')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
NewScreen 'frmUXGroupPolicy' 'Summary-group policy review' 11100;$f=$script:swiftForm;$f.ScrollBars=2
foreach($name in @('groupID','baseline')){[void](Control $f 109 $name 0 0 100 100);$script:swiftControl.Visible=$false}
[void](Control $f 100 'lblUIC' 400 1000 6000 350 'UIC / command grouping choice')
[void](Control $f 111 'cboUIC' 400 1400 12600 450);$script:swiftControl.RowSourceType='Value List';$script:swiftControl.LimitToList=$true
[void](Control $f 100 'lblBasis' 400 2000 12600 500 'Officer category basis: confirm Active or Reserve from personnel records; block 5 alone does not establish this.')
[void](Control $f 111 'cboBasis' 400 2600 12600 450);$script:swiftControl.RowSourceType='Value List';$script:swiftControl.RowSource='Unconfirmed;Active;Reserve';$script:swiftControl.LimitToList=$true
[void](Control $f 100 'lblNote' 400 3200 12600 500 'Review basis and UIC authority. Dual-hatted grouping requires the applicable flag-officer/SNDL authority; this app does not certify it.')
[void](Control $f 109 'txtNote' 400 3800 12600 1000);$script:swiftControl.ScrollBars=2
[void](Control $f 106 'chkComplete' 400 5050 300 300)
[void](Control $f 100 'lblComplete' 850 5000 12150 750 'I reviewed the complete eligible command roster for this group. This records a local statement; it does not waive findings. Recheck after saved-report or membership changes.')
[void](Control $f 109 'txtEvidence' 400 5900 12600 3300);$script:swiftControl.Locked=$true;$script:swiftControl.ScrollBars=2
[void](Control $f 104 'cmdSave' 400 9400 6100 500 'Record policy review' '=UXSaveGroupPolicy()')
[void](Control $f 104 'cmdBack' 6900 9400 6100 500 'Back to group' '=UXCloseScreen("frmUXGroupPolicy")')
[void](Control $f 100 'lblStatus' 400 10050 12600 800 '')
$f.OnLoad='=UXLoadGroupPolicy()';SaveScreen $f 'frmUXGroupPolicy'
$app.DoCmd.OpenForm('frmMajorWorkbench',1);$f=$app.Forms.Item('frmMajorWorkbench')
[void](Control $f 104 'uxPolicyReview' 8050 8050 5250 450 'Grouping policy and roster review' '=UXOpenGroupPolicy()')
SaveScreen $f 'frmMajorWorkbench'
$app.DoCmd.OpenForm('frmUXMembers',1);$f=$app.Forms.Item('frmUXMembers')
$f.Controls('cmdSave').Caption='Apply previewed membership';$f.Controls('cmdSave').OnClick='=UXApplyGroupMembers()'
$f.Controls('cmdSave').Top=9400;$f.Controls('cmdCancel').Top=9400;$f.Controls('lblStatus').Top=10000
[void](Control $f 104 'cmdPreviewImpact' 400 6250 12600 450 'Preview membership consequences' '=UXPreviewGroupMembers()')
[void](Control $f 109 'txtImpact' 400 6900 12600 2250);$script:swiftControl.Locked=$true;$script:swiftControl.ScrollBars=2
$f.Section(0).Height=10800;$f.ScrollBars=2;SaveScreen $f 'frmUXMembers'
$app.DoCmd.OpenForm('frmUXBulk',1);$f=$app.Forms.Item('frmUXBulk')
[void](Control $f 109 'txtGroupImpact' 2850 8800 10450 2900);$script:swiftControl.Locked=$true;$script:swiftControl.ScrollBars=2;$script:swiftControl.Visible=$false
$f.Section(0).Height=12000;SaveScreen $f 'frmUXBulk'
