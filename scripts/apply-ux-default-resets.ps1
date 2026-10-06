foreach($m in @('modUXGlobalDefaults','modUXBlockTools','modUX')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
$app.DoCmd.OpenForm('frmUXSharedFields',1);$f=$app.Forms.Item('frmUXSharedFields')
$keys=@($f.Controls | Where-Object {$_.Name.StartsWith('value')} | ForEach-Object {$_.Name.Substring(5)})
foreach($key in $keys){[void](Control $f 104 ('reset'+$key) 0 0 800 330 'Reset' ('=UXResetDefault("'+$key+'")'));$script:swiftControl.Visible=$false;$script:swiftControl.ControlTipText='Remove this explicit default override and inherit its parent; preview and apply first.'}
SaveScreen $f 'frmUXSharedFields'
