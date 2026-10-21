# Manual feedback phase 1; run after native workspace generation.
if(-not $script:swiftManagedBuild){foreach($m in @('modUXFeedback','modUX','modSwift','modSnippets','modUXSurface')){$app.LoadFromText(5,$m,(Join-Path $root ('src/'+$m+'.bas')))}}
$app.DoCmd.OpenForm('frmHome',1);$f=$app.Forms.Item('frmHome')
$f.Controls('lblHeading').Width=4500
[void](Control $f 100 'uxContributorLabel' 7800 900 5500 300 'Active contributor')
[void](Control $f 111 'uxContributor' 7800 1250 5500 420);$c=$script:swiftControl;$c.RowSourceType='Table/Query';$c.ColumnCount=2;$c.ColumnWidths='0;5000';$c.LimitToList=$true;$c.AfterUpdate='=UXOverviewContributor()'
[void](Control $f 100 'uxContributorStatus' 7800 1690 5500 300 '')
SaveScreen $f 'frmHome'
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
$f.Controls('cmdSnippets').OnClick='=UXOpenReportSnippets()'
[void](Control $f 104 'uxUndoPending' 10750 8200 2600 500 'Undo pending changes' '=UXUndoPending()')
SaveScreen $f 'frmEditor'
$app.DoCmd.OpenForm('frmSnippets',1);$f=$app.Forms.Item('frmSnippets');$f.Controls('cmdClose').OnClick='=UXReturnSnippets()';SaveScreen $f 'frmSnippets'
# Only this database's startup preference, not user-wide Access settings.
try{$app.CurrentDb().Properties('StartupShowDBWindow').Value=$false}catch{$db=$app.CurrentDb();$p=$db.CreateProperty('StartupShowDBWindow',1,$false);$db.Properties.Append($p)}
