param([string]$SourcePath=(Join-Path $PSScriptRoot '..\TEST_DB.accdb'))
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$SourcePath=[IO.Path]::GetFullPath($SourcePath)
$evidence=Join-Path $root 'planning\evidence\navfit'
[void](New-Item -ItemType Directory -Force -Path $evidence)
$before=(Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash
$scratch=Join-Path $evidence ('inspection-'+[guid]::NewGuid().ToString()+'.accdb')
Copy-Item -LiteralPath $SourcePath -Destination $scratch
$app=$null; $db=$null
try {
    $engine=New-Object -ComObject DAO.DBEngine.120
    Write-Host 'Created DAO engine'
    $db=$engine.OpenDatabase($scratch,$false,$true)
    Write-Host 'Opened read-only scratch'
    $tables=@()
    for($ti=0;$ti -lt $db.TableDefs.Count;$ti++) {
        $table=$db.TableDefs.Item($ti)
        Write-Host ('Inspecting table: '+$table.Name)
        $fields=@(); $indexes=@()
        for($fi=0;$fi -lt $table.Fields.Count;$fi++) {
            $field=$table.Fields.Item($fi)
            $properties=@{}
            foreach($key in @('Required','AllowZeroLength','DefaultValue','ValidationRule','Attributes','Description')) {
                try { $properties[$key]=$field.Properties.Item($key).Value } catch {}
            }
            $fields+=@{Name=$field.Name;Type=[int]$field.Type;Size=$field.Size;Properties=$properties}
        }
        for($ii=0;$ii -lt $table.Indexes.Count;$ii++) {
            $index=$table.Indexes.Item($ii)
            $indexFields=@(); for($ifi=0;$ifi -lt $index.Fields.Count;$ifi++) {$indexFields+=$index.Fields.Item($ifi).Name}
            $indexes+=@{Name=$index.Name;Primary=$index.Primary;Unique=$index.Unique;Fields=$indexFields}
        }
        $count=$null
        if(-not $table.Name.StartsWith('MSys')) { try { $rs=$db.OpenRecordset('SELECT Count(*) AS N FROM ['+$table.Name.Replace(']',']]')+']'); $count=$rs.Fields.Item('N').Value; $rs.Close() } catch {} }
        $tables+=@{Name=$table.Name;Attributes=$table.Attributes;RowCount=$count;Fields=$fields;Indexes=$indexes}
    }
    $relations=@()
    for($ri=0;$ri -lt $db.Relations.Count;$ri++) {
        $relation=$db.Relations.Item($ri)
        $relationFields=@();for($rfi=0;$rfi -lt $relation.Fields.Count;$rfi++){$rf=$relation.Fields.Item($rfi);$relationFields+=@{Name=$rf.Name;ForeignName=$rf.ForeignName}}
        $relations+=@{Name=$relation.Name;Table=$relation.Table;ForeignTable=$relation.ForeignTable;Attributes=$relation.Attributes;Fields=$relationFields}
    }
    $props=@{}
    for($pi=0;$pi -lt $db.Properties.Count;$pi++) { $p=$db.Properties.Item($pi);try { $props[$p.Name]=$p.Value } catch {} }
    $queries=@();for($qi=0;$qi -lt $db.QueryDefs.Count;$qi++){$q=$db.QueryDefs.Item($qi);$queries+=@{Name=$q.Name;Type=$q.Type;SQL=$q.SQL}}
    $result=@{SourceFile=[IO.Path]::GetFileName($SourcePath);SourceSHA256=$before;InspectedReadOnlyScratch=$scratch;EngineVersion=$db.Version;Properties=$props;Tables=$tables;Relations=$relations;Queries=$queries;Compatibility='Schema inspected only; NAVFIT reimport and printing untested'}
    $inventoryName=if([IO.Path]::GetFileName($SourcePath) -eq 'TEST_DB.accdb'){'schema-inventory.json'}else{[IO.Path]::GetFileNameWithoutExtension($SourcePath)+'.schema-inventory.json'}
    $result | ConvertTo-Json -Depth 16 | Set-Content -LiteralPath (Join-Path $evidence $inventoryName) -Encoding utf8
    $db.Close(); $db=$null
    if((Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash -ne $before) { throw 'Original source hash changed.' }
    $tables | Select-Object Name,RowCount,@{Name='Columns';Expression={$_.Fields.Count}} | Format-Table
    Write-Output "Original SHA256 unchanged: $before"
} finally {
    if($null -ne $db) {$db.Close()}
    if($null -ne $app) {$app.Quit(2); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
