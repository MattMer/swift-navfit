param([Parameter(Mandatory=$true)][string]$BuildPath,[string]$Variant)
#Before a schema change, capture the outgoing schema so the historical import
#matrix (test-swift-data-import.ps1) keeps testing it. Reads the build through
#read-only DAO and writes tests/historical-schemas/schema-NN[variant].sql.
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$dao=New-Object -ComObject DAO.DBEngine.120
$generated='^Index_[0-9A-F]{8}_[0-9A-F]{4}_[0-9A-F]{4}$'
$types=@{1='BIT';3='SHORT';4='LONG';7='DOUBLE';8='DATETIME';12='LONGTEXT'}
$db=$dao.OpenDatabase($BuildPath,$false,$true)
try{
 $r=$db.OpenRecordset('SELECT SchemaVersion,AppVersion FROM AppMetadata',4);$schema=[int]$r.Fields.Item(0).Value;$app=[string]$r.Fields.Item(1).Value;$r.Close()
 $tables=@(foreach($t in $db.TableDefs){if($t.Name -notlike 'MSys*' -and $t.Name -notlike '~*' -and -not $t.Connect){[string]$t.Name}})|Sort-Object
 $create=@();$indexes=@();$zeroLengthOff=@();$rules=@()
 foreach($name in $tables){
  $t=$db.TableDefs.Item($name);$columns=@()
  foreach($f in $t.Fields){
   if($f.DefaultValue -or $f.ValidationRule){throw ('Field properties are not captured yet: '+$name+'.'+$f.Name)}
   $type=[int]$f.Type
   if($f.Attributes -band 16){$columns+=('['+$f.Name+'] COUNTER');continue}
   $kind=if($type -eq 10){'TEXT('+$f.Size+')'}elseif($types.ContainsKey($type)){$types[$type]}else{throw ('Unsupported field type '+$type+': '+$name+'.'+$f.Name)}
   $columns+=('['+$f.Name+'] '+$kind+$(if($f.Required){' NOT NULL'}else{''}))
   if($type -in 10,12 -and -not $f.Required -and -not $f.AllowZeroLength){$zeroLengthOff+=($name+'.'+$f.Name)}
  }
  $create+=('CREATE TABLE ['+$name+'] ('+($columns -join ', ')+')')
  foreach($i in $t.Indexes){
   if($i.Foreign){continue}
   $indexName=if($i.Primary -and $i.Name -match $generated){'PK_'+$name}else{[string]$i.Name}
   $fields=(@(foreach($n in $i.Fields){'['+$n.Name+']'})-join ', ')
   $suffix=if($i.Primary){' WITH PRIMARY'}elseif($i.IgnoreNulls){' WITH IGNORE NULL'}else{''}
   $indexes+=('CREATE '+$(if($i.Unique -or $i.Primary){'UNIQUE '}else{''})+'INDEX ['+$indexName+'] ON ['+$name+'] ('+$fields+')'+$suffix)
  }
  if($t.ValidationRule){$rules+=('-- validationRule: '+$name+' '+$t.ValidationRule)}
 }
 $relations=@()
 foreach($rel in @($db.Relations|Sort-Object Name)){
  if($rel.Table -like 'MSys*'){continue}
  if($rel.Attributes -ne 0){throw ('Relation options are not captured yet: '+$rel.Name)}
  $fk=(@(foreach($f in $rel.Fields){'['+$f.ForeignName+']'})-join ', ');$pk=(@(foreach($f in $rel.Fields){'['+$f.Name+']'})-join ', ')
  $relations+=('ALTER TABLE ['+$rel.ForeignTable+'] ADD CONSTRAINT ['+$rel.Name+'] FOREIGN KEY ('+$fk+') REFERENCES ['+$rel.Table+'] ('+$pk+')')
 }
}finally{$db.Close()}
$folder=Join-Path $root 'tests/historical-schemas'
$existing=@(Get-ChildItem -LiteralPath $folder -Filter ('schema-{0:D2}*.sql' -f $schema))
if($existing.Count -and -not $Variant){throw ('Schema '+$schema+' already has '+(($existing|ForEach-Object Name)-join ', ')+'. Pass -Variant only if this build has a different structure.')}
$path=Join-Path $folder ('schema-{0:D2}{1}.sql' -f $schema,$Variant)
if(Test-Path -LiteralPath $path){throw ('Refusing to overwrite '+$path)}
$header=@('-- schema: '+$schema)
if($Variant){$header+=('-- variant: '+$Variant)}
$header+=@('-- kind: captured',('-- source: captured from '+[IO.Path]::GetFileName($BuildPath)+' ('+$app+')'),('-- appVersion: '+$app))
$header+=@($zeroLengthOff|ForEach-Object {'-- zeroLengthOff: '+$_})+$rules
[IO.File]::WriteAllText($path,((@($header)+$create+$indexes+$relations)-join "`n")+"`n",[Text.UTF8Encoding]::new($false))
Write-Output ('PASS: captured schema '+$schema+' ('+$app+'): '+$create.Count+' tables, '+$indexes.Count+' indexes, '+$relations.Count+' relations; '+$path)
