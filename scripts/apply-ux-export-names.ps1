foreach($m in @('modUXExportNames','modUXNAVFIT')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
$app.DoCmd.OpenForm('frmNAVFITExchange',1);$f=$app.Forms.Item('frmNAVFITExchange')
$f.Controls('lbltxtExportPath').Caption='Export directory'
$f.Controls('txtExportPath').Locked=$true;$f.Controls('txtExportPath').ControlTipText='Choose a directory. Swift assigns the rank and report ending date filename automatically when exporting.'
$f.Controls('cmdtxtExportPath').Caption='Choose folder'
SaveScreen $f 'frmNAVFITExchange'
