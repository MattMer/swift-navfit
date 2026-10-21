if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modReportNotes',(Join-Path $root 'src/modReportNotes.bas'))}
if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modSwift',(Join-Path $root 'src/modSwift.bas'))}
$db=$app.CurrentDb();$tables=@($db.TableDefs | ForEach-Object {$_.Name})
if(($tables -contains 'ReportNote') -xor ($tables -contains 'ReportNoteDraft')){throw 'Incomplete notes schema; refusing automatic repair'}
if($tables -notcontains 'ReportNote'){
 $schema=Get-Content (Join-Path $root 'src/major-schema.json') -Raw | ConvertFrom-Json
 $w=$app.DBEngine.Workspaces.Item(0);$w.BeginTrans()
 try{foreach($sql in $schema.SQL | Where-Object {$_ -match 'ReportNote|NoteDraft|LegacyNote'}){$db.Execute($sql,128)};$db.Execute('UPDATE AppMetadata SET SchemaVersion=22',128);$w.CommitTrans()}catch{$w.Rollback();throw}
}
$db.TableDefs.Refresh()
foreach($name in @('ReportNote','ReportNoteDraft')){$db.TableDefs.Item($name).Fields.Item('NoteText').AllowZeroLength=$true}
[void]$app.Eval('MigrateLegacyReportNotes()')
