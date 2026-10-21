if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modMemberHistory',(Join-Path $root 'src/modMemberHistory.bas'))}
if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modCommandDirectory',(Join-Path $root 'src/modCommandDirectory.bas'))}
if(-not $script:swiftManagedBuild){$app.LoadFromText(5,'modSwift',(Join-Path $root 'src/modSwift.bas'))}
$db=$app.CurrentDb();$tables=@($db.TableDefs | ForEach-Object {$_.Name})
$present=@('CommandMember','MemberReport','PriorEvaluation','PriorEvaluationTrait') | Where-Object {$tables -contains $_}
if($present.Count -gt 0 -and $present.Count -ne 4){throw 'Incomplete directory schema; refusing automatic repair'}
if($present.Count -eq 0){
 $schema=Get-Content (Join-Path $root 'src/major-schema.json') -Raw | ConvertFrom-Json
 $w=$app.DBEngine.Workspaces.Item(0);$w.BeginTrans()
 try{foreach($sql in $schema.SQL | Where-Object {$_ -match 'CommandMember|MemberReport|PriorEvaluation|UX_PriorGuid|PK_PriorTrait'}){$db.Execute($sql,128)};$db.Execute('UPDATE AppMetadata SET SchemaVersion=23',128);$w.CommitTrans()}catch{$w.Rollback();throw}
}
$db.TableDefs.Refresh()
foreach($name in @('CommandMember','PriorEvaluation','PriorEvaluationTrait')){foreach($field in $db.TableDefs.Item($name).Fields){if($field.Type -eq 10 -and -not $field.Required){$field.AllowZeroLength=$true}}}
[void]$app.Eval('SyncDirectoryFromReports()')
