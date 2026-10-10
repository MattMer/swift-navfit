param([string]$EditedPath,[string]$CandidatePath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent;$results=@()
$engine=New-Object -ComObject DAO.DBEngine.120
foreach($source in @($EditedPath,$CandidatePath)){
 $scratch=Join-Path $root ('build/ui-review-2026-10-07/data-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $source -Destination $scratch
 $db=$engine.OpenDatabase($scratch,$false,$true);$tables=@()
 try{foreach($t in $db.TableDefs){if($t.Name.StartsWith('MSys')){continue}
  $r=$db.OpenRecordset('SELECT * FROM ['+$t.Name+']');$rows=[Collections.Generic.List[string]]::new()
  while(-not $r.EOF){$row=@();foreach($field in $r.Fields){$v=$field.Value;if($v -is [datetime]){$v=$v.ToString('o')}elseif($v -is [DBNull]){$v=$null};$row+=,$v};$rows.Add(($row|ConvertTo-Json -Depth 8 -Compress));$r.MoveNext()};$r.Close()
  $payload=(@($rows|Sort-Object -CaseSensitive) -join "`n");$bytes=[Text.Encoding]::UTF8.GetBytes($payload);$hash=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
  $schema=@($t.Fields|ForEach-Object {@{name=$_.Name;type=$_.Type;size=$_.Size}})
  $tables+=@{name=$t.Name;rows=$rows.Count;sha256=$hash;schema=$schema}
 };$results+=@{file=[IO.Path]::GetFileName($source);tables=$tables}}finally{$db.Close()}
}
$left=@{};foreach($t in $results[0].tables){$left[$t.name]=$t};$differences=@()
foreach($t in $results[1].tables){$old=$left[$t.name];if($null -eq $old -or $old.sha256 -ne $t.sha256 -or ($old.schema|ConvertTo-Json -Depth 5 -Compress) -ne ($t.schema|ConvertTo-Json -Depth 5 -Compress)){$differences+=$t.name};$left.Remove($t.name)}
$differences+=@($left.Keys)
@{sources=$results;differences=$differences}|ConvertTo-Json -Depth 12|Set-Content (Join-Path $root 'planning/UI_AUTHOR_DATA_PRESERVATION_2026-10-07.json') -Encoding UTF8
if($differences.Count){throw ('Candidate table differences: '+($differences -join ', '))}
Write-Output ('PASS: all '+$results[0].tables.Count+' user tables retain identical schemas, row counts and values')
