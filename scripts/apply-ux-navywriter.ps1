. (Join-Path $PSScriptRoot 'access-properties.ps1')
if(-not $script:swiftManagedBuild){foreach($m in @('modNavyWriterCatalog','modUXNavyWriter')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
NewScreen 'frmUXNavyWriter' 'NavyWriter examples catalog' 11200;$f=$script:swiftForm;$f.ScrollBars=2
[void](Control $f 100 'lblScope' 400 950 12600 650 'Short examples with source links. Edit the wording for the member''s actual accomplishments before adding it to your snippets. Examples are not policy guidance.')
[void](Control $f 109 'txtSearch' 400 2000 6300 450)
[void](Control $f 100 'lblSearch' 400 1650 6300 300 'Search narrative sources')
[void](Control $f 100 'lblRate' 7200 1650 2700 300 'Rate / rating')
[void](Control $f 111 'cboRate' 7200 2000 2700 450);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;2500';$c.BoundColumn=1;$c.LimitToList=$true;$c.AfterUpdate='=UXLoadNavyWriter()'
[void](Control $f 104 'cmdSearch' 10300 2000 2700 450 'Search examples' '=UXLoadNavyWriter()')
[void](Control $f 110 'lstCatalog' 400 2500 12600 2200);$c=$script:swiftControl;$c.ColumnCount=4;$c.ColumnWidths='0;1300;2300;8500';$c.OnClick='=UXSelectNavyWriter()'
[void](Control $f 100 'lblBody' 400 4900 12600 350 'Preview and adapt the example')
[void](Control $f 109 'txtBody' 400 5350 12600 2200);$script:swiftControl.ScrollBars=2
[void](Control $f 109 'txtSource' 400 7800 12600 450);$script:swiftControl.Locked=$true
[void](Control $f 104 'cmdSource' 400 8550 3900 500 'Open full source' '=UXOpenNavyWriterSource()')
[void](Control $f 104 'cmdImport' 4750 8550 8250 500 'Add reviewed example to narrative snippets' '=UXImportNavyWriter()')
[void](Control $f 100 'lblStatus' 400 9300 12600 650 '')
[void](Control $f 104 'cmdBack' 400 10200 12600 500 'Back to Resources' '=UXCloseResource("frmUXNavyWriter")')
$f.OnLoad='=UXLoadNavyWriter()';SaveScreen $f 'frmUXNavyWriter'
$app.DoCmd.OpenForm('frmUXResources',1);$f=$app.Forms.Item('frmUXResources')
$bottom=([int](($f.Controls | Where-Object {$_.ControlType -eq 104 -and $_.Visible -and $_.Name -notmatch '^uxNav'} | ForEach-Object {$_.Top+$_.Height} | Measure-Object -Maximum).Maximum))+450
[void](Control $f 104 'uxNavyWriter' 2850 $bottom 10450 550 'NavyWriter examples catalog' '=OpenScreen("frmUXNavyWriter")')
SetAccessProperty ($f.Section(0)) 'Height' ([Math]::Max([int]$f.Section(0).Height,[int]($bottom+1100)));$f.ScrollBars=2
SaveScreen $f 'frmUXResources'
