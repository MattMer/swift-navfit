if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modSnippets',(Join-Path $root 'src/modSnippets.bas'))}
$app.DoCmd.OpenForm('frmSnippets',1);$f=$app.Forms.Item('frmSnippets')
$f.Controls('txtSearch').Width=5700;$f.Controls('cmdSearch').Left=10300;$f.Controls('cmdSearch').Width=2500
[void](Control $f 100 'lblRateFilter' 6400 1450 3500 300 'Filter by rate')
[void](Control $f 111 'cboRateFilter' 6400 1800 3500 450);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;3200';$c.LimitToList=$true;$c.AfterUpdate='=LoadSnippets()'
[void](Control $f 100 'lblSnippetRate' 6800 6200 1500 350 'Snippet rate')
[void](Control $f 111 'cboSnippetRate' 8500 6200 4300 450);$c=$script:swiftControl;$c.ColumnCount=2;$c.ColumnWidths='0;4000';$c.LimitToList=$true
SaveScreen $f 'frmSnippets'
