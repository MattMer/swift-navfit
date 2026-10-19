foreach($m in @('modUXReadRows','modUXFeedback','modLibraryViews','modUX','modUXDashboard')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
$app.DoCmd.OpenForm('frmUXGroups',1);$f=$app.Forms.Item('frmUXGroups')
[void](Control $f 104 'uxEditDefaults' 2850 7450 10450 450 'Edit selected group defaults' '=UXListGroupDefaults()')
$f.Controls('uxLifecycleStatus').Top=8050;$f.Section(0).Height=8750
SaveScreen $f 'frmUXGroups'
