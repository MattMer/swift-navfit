if(-not $script:swiftManagedBuild){foreach($m in @('modUXGlobalDefaults','modUXBlockTools','modUX','modSwift')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
$app.DoCmd.OpenForm('frmUXGroups',1);$f=$app.Forms.Item('frmUXGroups')
[void](Control $f 104 'uxGlobalDefaults' 2850 9850 10450 450 'Edit global defaults for all summary groups' '=UXOpenGlobalDefaults()')
$f.Section(0).Height=10700;$f.ScrollBars=2
SaveScreen $f 'frmUXGroups'
