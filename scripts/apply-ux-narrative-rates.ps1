foreach($m in @('modNavyWriterCatalog','modUXNavyWriter')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}
$app.DoCmd.OpenForm('frmUXNavyWriter',1);$f=$app.Forms.Item('frmUXNavyWriter')
$f.Controls.Item('txtSearch').Width=6300;$f.Controls.Item('txtSearch').Top=2000;$f.Controls.Item('cmdSearch').Top=2000
[void](Control $f 100 'lblSearch' 400 1650 6300 300 'Search narrative sources')
[void](Control $f 100 'lblRate' 7200 1650 2700 300 'Rate / rating')
[void](Control $f 111 'cboRate' 7200 2000 2700 450);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;2500';$c.BoundColumn=1;$c.LimitToList=$true;$c.AfterUpdate='=UXLoadNavyWriter()'
$f.Controls.Item('lstCatalog').ColumnCount=4;$f.Controls.Item('lstCatalog').ColumnWidths='0;1300;2300;8500'
SaveScreen $f 'frmUXNavyWriter'
