param([Parameter(Mandatory=$true)][string]$SourcePath,[Parameter(Mandatory=$true)][string]$CandidatePath,[switch]$AllowDemoSeed)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$before=@((Get-FileHash -LiteralPath $SourcePath).Hash,(Get-FileHash -LiteralPath $CandidatePath).Hash)
$folder=Join-Path $root ('build/tests/group-directory-data-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
Copy-Item -LiteralPath $SourcePath -Destination (Join-Path $folder 'source.accdb')
Copy-Item -LiteralPath $CandidatePath -Destination (Join-Path $folder 'candidate.accdb')
$engine=New-Object -ComObject DAO.DBEngine.120
$left=$engine.OpenDatabase((Join-Path $folder 'source.accdb'),$false,$true)
$right=$engine.OpenDatabase((Join-Path $folder 'candidate.accdb'),$false,$true)
$tables=@();$rows=0;$differences=@()
try{
 foreach($table in $left.TableDefs){
  if($table.Name.StartsWith('MSys')){continue}
  $name=$table.Name;$target=$right.TableDefs.Item($name);$where='';$maxID=0
  if($AllowDemoSeed -and $name -in @('CommandMember','OperationJournal')){$idField=$table.Fields.Item(0).Name;$m=$left.OpenRecordset('SELECT Max(['+$idField+']) FROM ['+$name+']');if($m.Fields.Item(0).Value -isnot [DBNull]){$maxID=[int]$m.Fields.Item(0).Value};$m.Close();$where=' WHERE ['+$idField+']<='+$maxID}
  $fields=@($table.Fields|ForEach-Object {$_.Name})
  foreach($field in $table.Fields){$other=$target.Fields.Item($field.Name);if($other.Type -ne $field.Type -or $other.Size -ne $field.Size){throw ('Existing field schema changed: '+$name+'.'+$field.Name)}}
  $extra=@($target.Fields|Where-Object {$fields -notcontains $_.Name}|ForEach-Object {$_.Name})
  if($extra.Count -and ($name -ne 'CommandMember' -or $extra.Count -ne 1 -or $extra[0] -ne 'RUADData')){throw ('Unexpected added field: '+$name)}
  $columns=($fields|ForEach-Object {'['+$_+']'}) -join ','
  $payloads=@()
  foreach($db in @($left,$right)){
   $r=$db.OpenRecordset('SELECT '+$columns+' FROM ['+$name+']'+$where);$records=[Collections.Generic.List[string]]::new()
   while(-not $r.EOF){
    $values=@();foreach($field in $r.Fields){$v=$field.Value;if($v -is [datetime]){$v=$v.ToString('o')}elseif($v -is [DBNull]){$v=$null};if($name -eq 'AppMetadata' -and $field.Name -eq 'SchemaVersion'){$v='documented additive migration'};$values+=,$v}
    $records.Add(($values|ConvertTo-Json -Depth 8 -Compress));$r.MoveNext()
   };$r.Close()
   $text=@($records|Sort-Object -CaseSensitive) -join "`n"
   $sha=[Security.Cryptography.SHA256]::Create();$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text)))).Replace('-','');$sha.Dispose()
   $payloads+=@{rows=$records.Count;hash=$hash}
  }
  if($payloads[0].rows -ne $payloads[1].rows -or $payloads[0].hash -ne $payloads[1].hash){$differences+=$name}
  $rows+=$payloads[0].rows;$tables+=@{name=$name;rows=$payloads[0].rows;sha256=$payloads[0].hash;newFields=$extra}
 }
 if(@($right.TableDefs|Where-Object {-not $_.Name.StartsWith('MSys')}).Count -ne $tables.Count){throw 'Unexpected table count'}
 if($differences.Count){throw ('Existing data changed in: '+($differences -join ', '))}
 $r=$right.OpenRecordset('SELECT SchemaVersion FROM AppMetadata');if($r.Fields.Item(0).Value -ne 25){throw 'Candidate schema must be 25'};$r.Close()
 $memberMax=$left.OpenRecordset('SELECT Max(MemberID) FROM CommandMember');$memberLimit=[int]$memberMax.Fields.Item(0).Value;$memberMax.Close()
 if($AllowDemoSeed){$seed=$right.OpenRecordset("SELECT Count(*) FROM CommandMember WHERE SourceKind='Demo RUAD' AND MemberID>"+$memberLimit);if($seed.Fields.Item(0).Value -ne 2){throw 'Expected two synthetic RUAD demo profiles'};$seed.Close()}
 $r=$right.OpenRecordset('SELECT Count(*) FROM CommandMember WHERE RUADData Is Not Null AND MemberID<='+$memberLimit);if($r.Fields.Item(0).Value -ne 0){throw 'New RUAD values unexpectedly altered existing profiles'};$r.Close()
}finally{$left.Close();$right.Close()}
if((Get-FileHash -LiteralPath $SourcePath).Hash -ne $before[0] -or (Get-FileHash -LiteralPath $CandidatePath).Hash -ne $before[1]){throw 'Source/candidate file changed'}
$evidence=@{status='PASS';source=[IO.Path]::GetFileName($SourcePath);candidate=[IO.Path]::GetFileName($CandidatePath);tables=$tables;rows=$rows;sourceHash=$before[0];candidateHash=$before[1];allowedChanges=@('AppMetadata.SchemaVersion 24 to 25','CommandMember.RUADData nullable LONGTEXT');sourceAndCandidateUnchanged=$true;twoSyntheticRUADProfilesAdded=[bool]$AllowDemoSeed}
$evidence|ConvertTo-Json -Depth 8|Set-Content (Join-Path $root 'planning/GROUP_DIRECTORY_DATA_PRESERVATION_2026-10-07.json') -Encoding utf8
Write-Output ('PASS: '+$tables.Count+' tables and '+$rows+' existing rows preserved; only documented additive schema-25 changes')
