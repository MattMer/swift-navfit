param([Parameter(Mandatory=$true)][string]$BuildPath,[string]$Schemas,[string]$CurrentSchemaPath,[switch]$CurrentFixturesOnly,[switch]$SkipKeptReleases)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$folder=Join-Path $root ('build/tests/data-import-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
$dao=New-Object -ComObject DAO.DBEngine.120
#Every schema from 1 to the build's schema gets a populated synthetic source.
#tests/historical-schemas holds one definition file per older schema (exact git
#DDL, DDL captured from a real build, or a documented reconstruction); the
#current schema is a copy of the build itself. Synthetic rows exercise every
#table and column, and this runner checks the transferred values itself instead
#of relying only on the importer's own verification.
function ReadDefinition($file){
 $entry=@{sql=@();notes=@();zeroLengthOff=@();tableValidationRules=@{};file=$file.Name}
 foreach($line in [IO.File]::ReadAllLines($file.FullName,[Text.Encoding]::UTF8)){
  if($line -match '^-- (\w+): (.*)$'){
   switch($Matches[1]){
    'schema'{$entry.schema=[int]$Matches[2]}
    'note'{$entry.notes+=$Matches[2]}
    'zeroLengthOff'{$entry.zeroLengthOff+=$Matches[2]}
    'validationRule'{$parts=$Matches[2].Split(' ',2);$entry.tableValidationRules[$parts[0]]=$parts[1]}
    default{$entry[$Matches[1]]=$Matches[2]}
   }
  }elseif($line.Trim()){$entry.sql+=$line}
 }
 return [pscustomobject]$entry
}
$catalogue=@(Get-ChildItem -LiteralPath (Join-Path $root 'tests/historical-schemas') -Filter 'schema-*.sql'|Sort-Object Name|ForEach-Object {ReadDefinition $_})
#Non-ASCII text built from code points keeps this file ASCII for Windows PowerShell.
$unicode=-join @([char]0x00C9,[char]0x00FC,[char]0x4E2D,[char]0x2019)
$fieldDefinitions=@{};foreach($f in (Get-Content (Join-Path $root 'src/field-definitions.json') -Raw -Encoding UTF8|ConvertFrom-Json).fields){$fieldDefinitions[$f.key]=$f}
function Scalar([string]$path,[string]$sql){$db=$dao.OpenDatabase($path,$false,$true);try{$r=$db.OpenRecordset($sql,4);$v=$r.Fields.Item(0).Value;$r.Close();return $v}finally{$db.Close()}}
$current=[int](Scalar $BuildPath 'SELECT SchemaVersion FROM AppMetadata')
#-Schemas takes a list such as '1-7,20' for focused runs; the default is every schema.
$requested=if($Schemas){@($Schemas -split ','|ForEach-Object {if($_ -match '^\s*(\d+)\s*-\s*(\d+)\s*$'){[int]$Matches[1]..[int]$Matches[2]}else{[int]$_}})}elseif($CurrentFixturesOnly){@($current)}else{@(1..$current)}
# Reflection preserves DAO's Variant setter across mixed string/date/number fields.
function SetField($record,[string]$key,$value){$field=$record.Fields.Item($key);[void]$field.GetType().InvokeMember('Value',[Reflection.BindingFlags]::SetProperty,$null,$field,@($value))}
function IsNullValue($value){return $null -eq $value -or [Convert]::IsDBNull($value)}
function UserTables($db){return @($db.TableDefs|Where-Object {$_.Name -notlike 'MSys*' -and $_.Name -notlike '~*' -and -not $_.Connect}|ForEach-Object {$_.Name})}

#Every historical build set AllowZeroLength on optional text; zeroLengthOff lists exceptions.
function ApplyFieldProperties($db,$entry){
 foreach($name in (UserTables $db)){foreach($field in @($db.TableDefs.Item($name).Fields)){if($field.Type -in 10,12 -and -not $field.Required){$field.AllowZeroLength=$true}}}
 foreach($item in @($entry.zeroLengthOff)){if($item){$parts=$item.Split('.');$db.TableDefs.Item($parts[0]).Fields.Item($parts[1]).AllowZeroLength=$false}}
 foreach($rule in $entry.tableValidationRules.GetEnumerator()){$db.TableDefs.Item($rule.Key).ValidationRule=$rule.Value}
}
#A schema can have structural variants when builds changed without a version bump.
function NewFixture([int]$schema,$entry,[string]$path){
 if(-not $entry){Copy-Item -LiteralPath $BuildPath -Destination $path;return @{Kind='current-build';Definition=$BuildPath}}
 $db=$dao.CreateDatabase($path,';LANGID=0x0409;CP=1252;COUNTRY=0',128)
 try{
  foreach($sql in $entry.sql){try{$db.Execute($sql,128)}catch{throw ('Schema '+$schema+' DDL failed: '+$sql+"`n"+$_.Exception.Message)}}
  ApplyFieldProperties $db $entry
  if(-not @($entry.sql|Where-Object {$_ -like 'INSERT INTO AppMetadata*'}).Count){$db.Execute("INSERT INTO AppMetadata (MetadataID,AppVersion,SchemaVersion) VALUES (1,'"+$entry.appVersion+"',"+$schema+")",128)}
 }finally{$db.Close()}
 if([int](Scalar $path 'SELECT SchemaVersion FROM AppMetadata') -ne $schema){throw ('Fixture for schema '+$schema+' has the wrong SchemaVersion')}
 return @{Kind=$entry.kind;Definition=($entry.file+': '+$entry.source)}
}

#Tables derived by the importer itself, or reference rows the build already owns.
$derivedTables=@('AppMetadata','TraitDefinition','RevisionDate','RevisionChoice','GroupCompletion','GroupCompletionReport','GroupCompletionCurrent','RSCountedReport','FieldDefaultLink','GroupLifecycle')
function TextValue($field,[string]$table,[int]$i,$definition){
 $size=[int]$field.Size;$name=$field.Name
 if($definition -and $definition.options -and $definition.options.Count){return [string]$definition.options[$i % $definition.options.Count]}
 if($definition -and $definition.date){return ('2025-0'+($i+1)+'-1'+$i)}
 if($name -like '*Guid'){return ([guid]::NewGuid().ToString('B').ToUpperInvariant())}
 if($field.Type -eq 12){return ("Synthetic $table.$name row $i`r`nLine two keeps CRLF, Unicode $unicode and 'quotes'.`r`n"+('Long memo text. '*40)+'Trailing space ')}
 $text=switch($i % 3){0{"Synthetic $name $i"}1{"$unicode '$i' "}default{"syn $name $i"}}
 if($text.Length -le $size){return $text}
 $short=('S'+$i+$name)
 return $short.Substring(0,[Math]::Min($size,$short.Length))
}
#Coded columns get values the application of that schema could have written.
$domains=@{
 'SummaryGroup.KeySpec'={param($i,$schema) if($schema -lt 23){[DBNull]::Value}else{'<SwiftGroupPolicy version="1"/>'}}
 'SummaryGroup.FormKind'={'EVAL'}
 'SummaryGroup.Paygrade'={'E5'}
 'GroupMember.Paygrade'={'E5'}
 'Report.WorkflowState'={param($i,$schema) if($schema -lt 20){'Draft'}else{@('Draft','Approved')[$i % 2]}}
 'Report.LegacyFormKind'={[DBNull]::Value}
 'Report.LegacyWorkflowState'={[DBNull]::Value}
 'Report.LegacyWorkflowNull'={$false}
 #Audit actors mirror the contributor they register (ActorID is the UserProfileID).
 'AuditActor.ActorID'={param($i) (ActorProfile $i).UserProfileID}
 'AuditActor.ProfileGuid'={param($i) (ActorProfile $i).ProfileGuid}
 'AuditActor.RegisteredName'={param($i) (ActorProfile $i).DisplayName}
 'AuditEvent.SourceKind'={'Revision'}
 'AuditEvent.ReportEventSourceID'={[DBNull]::Value}
 'AuditEvent.FolderMoveSourceID'={[DBNull]::Value}
 'AuditEvent.CommentEventSourceID'={[DBNull]::Value}
 'AuditEvent.TaskEventSourceID'={[DBNull]::Value}
 'AuditEvent.JournalSourceID'={[DBNull]::Value}
 'AuditEvent.MessageSourceID'={[DBNull]::Value}
 'ReportSignoff.PriorState'={param($i) @('Draft','Approved')[$i % 2]}
 'ReportSignoff.ReopenedBy'={[DBNull]::Value}
 'ReportSignoff.ReopenedAt'={[DBNull]::Value}
 'ReportSignoff.ReopenReason'={[DBNull]::Value}
 'MemberCounseling.Outcome'={param($i) @('Completed','Not performed')[$i % 2]}
 'MemberCounseling.CounselorName'={param($i) if($i % 2){[DBNull]::Value}else{'EXAMPLE, C'}}
 'MemberCounseling.NotPerformedReason'={param($i) if($i % 2){'TEMADD'}else{[DBNull]::Value}}
 'MemberCounseling.WorksheetDisposition'={param($i) if($i % 3 -eq 0){'Given to member'}else{[DBNull]::Value}}
 'MemberCounseling.DispositionOn'={param($i) if($i % 3 -eq 0){[datetime]'2026-10-20'}else{[DBNull]::Value}}
}
#Tables that copy identity from another table without a relationship to it.
$seedAfter=@{AuditActor='UserProfile'}
function ActorProfile([int]$i){$profiles=@($rows['UserProfile']);return $profiles[$i % $profiles.Count].Values}
function SeedValue($field,[string]$table,[int]$i,[int]$schema){
 $definition=if($table -in 'Report','RecoveryDraft'){$fieldDefinitions[$field.Name]}else{$null}
 switch([int]$field.Type){
  1{return ($i % 2 -eq 0)}
  {$_ -in 2,3,4}{return ($i+1)}
  5{return ([decimal]12.34+$i)}
  6{return ([single]1.5+$i)}
  7{return ([double](1.0/3.0)+$i)}
  8{return ([datetime]'2024-02-29 13:45:10').AddDays($i)}
  {$_ -in 10,12}{return (TextValue $field $table $i $definition)}
  15{return ([guid]::NewGuid().ToString('B'))}
  default{throw ('Unsupported seed field type '+$table+'.'+$field.Name+' ('+$field.Type+')')}
 }
}
function KeyOf($record,$keyFields){return ((@($keyFields|ForEach-Object {[string]$record.Fields.Item($_).Value}))-join '|')}
function PrimaryFields($table){foreach($index in @($table.Indexes)){if($index.Primary){return @($index.Fields|ForEach-Object {$_.Name})}};return @()}

function SeedFixture([string]$path,[int]$schema){
 $db=$dao.OpenDatabase($path)
 try{
  $names=@(UserTables $db|Where-Object {$_ -notin $derivedTables})
  $parents=@{}
  foreach($relation in @($db.Relations)){if($relation.Fields.Count -eq 1){$f=$relation.Fields.Item(0);$parents[$relation.ForeignTable+'.'+$f.ForeignName]=@{Table=$relation.Table;Field=$f.Name}}}
  $rows=@{};$deferred=@();$done=@{};$order=@()
  #References order the tables; an optional reference that closes a cycle is deferred.
  function Ready([string]$table,[bool]$requiredOnly){
   if($seedAfter[$table] -and $names -contains $seedAfter[$table] -and -not $done[$seedAfter[$table]]){return $false}
   $key=PrimaryFields $db.TableDefs.Item($table)
   foreach($f in @($db.TableDefs.Item($table).Fields)){$p=$parents[$table+'.'+$f.Name];if($p -and $p.Table -ne $table -and ($f.Required -or $key -contains $f.Name -or -not $requiredOnly) -and $names -contains $p.Table -and -not $done[$p.Table]){return $false}}
   return $true
  }
  while($order.Count -lt $names.Count){
   $pending=@($names|Where-Object {-not $done[$_]})
   $next=$pending|Where-Object {Ready $_ $false}|Select-Object -First 1
   if(-not $next){$next=$pending|Where-Object {Ready $_ $true}|Select-Object -First 1}
   if(-not $next){throw ('Required reference cycle among '+(($names|Where-Object {-not $done[$_]})-join ', '))}
   $done[$next]=$true;$order+=$next
  }
  $reportFields=@(if($names -contains 'Report'){@($db.TableDefs.Item('Report').Fields|ForEach-Object {$_.Name})})
  foreach($name in $order){
   if($name -in 'RevisionValue','ReportHead'){continue}
   $t=$db.TableDefs.Item($name);$keys=PrimaryFields $t
   $existing=[int]$db.OpenRecordset('SELECT Count(*) FROM ['+$name+']',4).Fields.Item(0).Value
   if($name -eq 'Workspace' -and $existing -gt 0){continue}
   $selfReference=@($t.Fields|Where-Object {$p=$parents[$name+'.'+$_.Name];$p -and $p.Table -eq $name}).Count -gt 0
   $count=if($name -eq 'Workspace'){1}elseif($selfReference -or $name -eq 'ReportRevision'){6}else{3}
   $rows[$name]=@();$picked=@{}
   $uniqueFields=@(foreach($index in @($t.Indexes)){if($index.Unique -and -not $index.Primary -and $index.Fields.Count -eq 1){$index.Fields.Item(0).Name}})
   $fields=@($t.Fields|Sort-Object {if($_.Name -eq 'ReportID'){0}else{1}})
   for($i=0;$i -lt $count;$i++){
    $r=$db.OpenRecordset($name,2);$r.AddNew();$report=$null;$late=@()
    foreach($f in $fields){
     if($f.Attributes -band 16){continue}
     $domain=$domains[$name+'.'+$f.Name];if($domain){SetField $r $f.Name (& $domain ($i+$existing) $schema);continue}
     $p=$parents[$name+'.'+$f.Name]
     if($name -eq 'Workspace' -and $f.Name -eq 'WorkspaceID'){SetField $r $f.Name 1;continue}
     if($p -and $names -notcontains $p.Table){if($f.Required){throw ('Cannot seed '+$name+'.'+$f.Name+': parent '+$p.Table+' is derived')};continue}
     if($p){
      $candidates=@($rows[$p.Table])
      if($p.Table -eq $name){$candidates=@($candidates|Where-Object {$null -eq $report -or $_.Report -eq $report});if($candidates.Count){SetField $r $f.Name $candidates[-1].Values[$p.Field]};continue}
      if(-not $rows.ContainsKey($p.Table)){$late+=@{Field=$f.Name;Parent=$p};continue}
      if($null -ne $report -and $p.Table -ne 'Report'){$same=@($candidates|Where-Object {$_.Report -eq $report});if($same.Count){$candidates=$same}}
      if(-not $candidates.Count){if($f.Required){throw ('No parent rows for '+$name+'.'+$f.Name)};continue}
      $value=$candidates[$i % $candidates.Count].Values[$p.Field]
      if($uniqueFields -contains $f.Name){
       if(-not $picked[$f.Name]){$picked[$f.Name]=@{}}
       $free=@($candidates|Where-Object {-not $picked[$f.Name].ContainsKey([string]$_.Values[$p.Field])})
       if(-not $free.Count){if($f.Required){throw ('Not enough parents for unique '+$name+'.'+$f.Name)};continue}
       $value=$free[0].Values[$p.Field];$picked[$f.Name][[string]$value]=$true
      }
      SetField $r $f.Name $value
      if($f.Name -eq 'ReportID'){$report=$value}
      continue
     }
     if($i -eq 2 -and -not $f.Required -and $keys -notcontains $f.Name){continue}
     if($i -eq 1 -and $f.Type -in 10,12 -and $f.AllowZeroLength -and ($f.OrdinalPosition % 2)){SetField $r $f.Name '';continue}
     $value=SeedValue $f $name ($i+$existing) $schema
     if($f.Name -eq 'ReportID' -and $name -ne 'Report'){$report=$value}
     SetField $r $f.Name $value
    }
    if($name -eq 'ReportRevision'){
     $number=@($rows[$name]|Where-Object {$_.Report -eq $report}).Count+1;SetField $r 'RevisionNumber' $number
     if($t.Fields|Where-Object {$_.Name -eq 'SnapshotSchema'}){SetField $r 'SnapshotSchema' $schema}
     if($t.Fields|Where-Object {$_.Name -eq 'FieldDefinitionVersion'}){SetField $r 'FieldDefinitionVersion' $(if(@($fieldDefinitions.Keys|Where-Object {$reportFields -notcontains $_}).Count){[DBNull]::Value}else{1})}
    }
    try{$r.Update()}catch{throw ("Seeding schema "+$schema+" "+$name+" row "+$i+" failed: "+$_.Exception.Message)};$r.Bookmark=$r.LastModified
    $values=@{};foreach($f in @($t.Fields)){$values[$f.Name]=$r.Fields.Item($f.Name).Value}
    if($name -eq 'Report'){$report=$values['ReportID']}
    $rows[$name]+=@{Values=$values;Report=$report;Key=(KeyOf $r $keys)}
    foreach($item in $late){$deferred+=@{Table=$name;Keys=$keys;Values=$values;Report=$report;Field=$item.Field;Parent=$item.Parent}}
    $r.Close()
   }
  }
  #Each deferred reference takes a parent no other row uses, so unique links stay unique.
  $used=@{}
  foreach($item in $deferred){
   if($item.Table -eq 'Report' -and $item.Field -eq 'CurrentRevisionID'){continue}
   $candidates=@($rows[$item.Parent.Table]);$same=@($candidates|Where-Object {$_.Report -eq $item.Report});if($same.Count){$candidates=$same}
   $slot=$item.Table+'.'+$item.Field;if(-not $used[$slot]){$used[$slot]=@{}}
   $parent=@($candidates|Where-Object {-not $used[$slot].ContainsKey([string]$_.Values[$item.Parent.Field])})|Select-Object -First 1
   if(-not $parent){continue}
   $used[$slot][[string]$parent.Values[$item.Parent.Field]]=$true
   $where=(@($item.Keys|ForEach-Object {'['+$_+']='+$item.Values[$_]})-join ' AND ')
   $sql='UPDATE ['+$item.Table+'] SET ['+$item.Field+']='+$parent.Values[$item.Parent.Field]+' WHERE '+$where;try{$db.Execute($sql,128)}catch{throw ('Deferred reference failed: '+$sql+' '+$_.Exception.Message)}
  }
  #A report's saved revision owns its values: the current revision's snapshot
  #equals the Report columns, as every Swift version wrote it.
  if($rows.ContainsKey('Report') -and $rows.ContainsKey('ReportRevision')){
   $value=$db.OpenRecordset('RevisionValue',2)
   foreach($report in $rows['Report']){
    $revisions=@($rows['ReportRevision']|Where-Object {$_.Report -eq $report.Report})
    if(-not $revisions.Count){throw ('Seeded report '+$report.Report+' has no revision')}
    $head=$revisions[-1].Values['RevisionID']
    $db.Execute('UPDATE Report SET CurrentRevisionID='+$head+' WHERE ReportID='+$report.Report,128)
    $snapshot=$db.OpenRecordset('SELECT * FROM Report WHERE ReportID='+$report.Report,4)
    foreach($revision in $revisions){
     foreach($key in $reportFields){
      if(-not $fieldDefinitions.ContainsKey($key)){continue}
      $v=$snapshot.Fields.Item($key).Value
      if($revision.Values['RevisionID'] -ne $head -and $key -eq 'Title'){$v='Earlier saved title'}
      $value.AddNew();SetField $value 'RevisionID' $revision.Values['RevisionID'];SetField $value 'FieldKey' $key
      if(IsNullValue $v){SetField $value 'IsNullValue' $true}else{SetField $value 'IsNullValue' $false;if($v -is [datetime]){$v=$v.ToString('yyyy-MM-dd')};SetField $value 'ValueText' ([string]$v)}
      $value.Update()
     }
    }
    $snapshot.Close()
    if($names -contains 'ReportHead'){$db.Execute('INSERT INTO ReportHead (ReportID,RevisionID) VALUES ('+$report.Report+','+$head+')',128)}
   }
   $value.Close()
  }
  $total=0;foreach($name in (UserTables $db)){if($name -ne 'AppMetadata'){$total+=[int]$db.OpenRecordset('SELECT Count(*) FROM ['+$name+']',4).Fields.Item(0).Value}}
  return $total
 }finally{$db.Close()}
}

#Independent exact comparison: every source row must exist in the target with
#identical values (binary text, null versus empty, dates and numbers).
function CompareTransfer([string]$source,[string]$target){
 $s=$dao.OpenDatabase($source,$false,$true);$t=$dao.OpenDatabase($target,$false,$true);$checked=0
 try{
  foreach($name in (UserTables $s)){
   if($name -eq 'AppMetadata'){continue}
   $table=$s.TableDefs.Item($name);$keys=PrimaryFields $table
   if(-not $keys.Count){throw ('Source table has no primary key: '+$name)}
   $fieldNames=@($table.Fields|ForEach-Object {$_.Name})
   $list=(@($fieldNames|ForEach-Object {'['+$_+']'})-join ',')
   $targetRows=@{};$r=$t.OpenRecordset('SELECT '+$list+' FROM ['+$name+']',4)
   while(-not $r.EOF){$values=@{};foreach($n in $fieldNames){$values[$n]=$r.Fields.Item($n).Value};$targetRows[(KeyOf $r $keys)]=$values;$r.MoveNext()}
   $r.Close();$r=$s.OpenRecordset('SELECT '+$list+' FROM ['+$name+']',4)
   while(-not $r.EOF){
    $key=KeyOf $r $keys;$other=$targetRows[$key]
    if(-not $other){throw ('Target is missing '+$name+' row '+$key)}
    foreach($n in $fieldNames){
     $a=$r.Fields.Item($n).Value;$b=$other[$n]
     $same=if((IsNullValue $a) -or (IsNullValue $b)){(IsNullValue $a) -and (IsNullValue $b)}elseif($a -is [string]){[string]::Equals($a,[string]$b,[StringComparison]::Ordinal)}else{[object]::Equals($a,$b)}
     if(-not $same){throw ('Value differs in '+$name+' row '+$key+' field '+$n)}
    }
    $checked++;$r.MoveNext()
   }
   $r.Close()
  }
 }finally{$s.Close();$t.Close()}
 return $checked
}

$fixtures=[ordered]@{}
foreach($schema in ($requested|Sort-Object -Unique)){
 if($schema -lt 1 -or $schema -gt $current){throw ('Schema '+$schema+' is outside 1-'+$current)}
 $entries=@(if($schema -eq $current){$null}else{$catalogue|Where-Object {$_.schema -eq $schema}})
 if(-not $entries.Count){throw ('tests/historical-schemas has no definition for schema '+$schema+'; capture it with scripts/capture-schema-definition.ps1')}
 foreach($entry in $entries){
  $label=[string]$schema;if($entry -and $entry.variant){$label+=$entry.variant}
  $path=Join-Path $folder ('schema-'+$label+'-source.accdb')
  $made=NewFixture $schema $entry $path
  $made.Rows=SeedFixture $path $schema
  $fixtures[$label]=@{Schema=$schema;Variant=$label;Kind=$made.Kind;Definition=$made.Definition;Source=$path;Rows=$made.Rows}
 }
}
#Kept populated releases add real application-written data on top of the matrix.
$kept=@()
if(-not $CurrentFixturesOnly -and -not $SkipKeptReleases){
 $best=@{}
 foreach($file in Get-ChildItem -LiteralPath (Join-Path $root 'output') -Recurse -File|Where-Object {$_.Extension -eq '.accdb' -and $_.Name -notmatch '^upgrade-| - Copy'}){
  try{$v=[int](Scalar $file.FullName 'SELECT SchemaVersion FROM AppMetadata');$n=[int](Scalar $file.FullName 'SELECT Count(*) FROM Report')}catch{continue}
  if($n -gt 0 -and $requested -contains $v -and (-not $best.ContainsKey($v) -or $file.LastWriteTime -gt $best[$v].LastWriteTime)){$best[$v]=$file}
 }
 foreach($v in ($best.Keys|Sort-Object)){
  $copy=Join-Path $folder ('kept-schema-'+$v+'.accdb');Copy-Item -LiteralPath $best[$v].FullName -Destination $copy
  $kept+=@{Schema=$v;Kind='kept-release';Definition=$best[$v].FullName;Source=$copy;OriginalSHA256=(Get-FileHash -LiteralPath $best[$v].FullName).Hash}
 }
}
if($CurrentSchemaPath){$kept+=@{Schema=[int](Scalar $CurrentSchemaPath 'SELECT SchemaVersion FROM AppMetadata');Kind='supplied';Definition=$CurrentSchemaPath;Source=$CurrentSchemaPath}}

$results=@();$keptResults=@();$app=New-Object -ComObject Access.Application
function OpenTarget([string]$name){
 $target=Join-Path $folder ($name+'-target.accdb');Copy-Item -LiteralPath $BuildPath -Destination $target
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
 foreach($form in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$form,2)}
 $app.LoadFromText(5,'WorkspaceImportHarness',(Join-Path $root 'tests/WorkspaceImportHarness.bas'))
 [void]$app.Eval('ConfigureWorkspace("SYNTHETIC IMPORT BOOTSTRAP","Tester")')
 return $target
}
function CallImport([string]$source){return $app.Eval('TryWorkspaceImport("'+$source.Replace('"','""')+'")')}
function ImportOne($fixture,[string]$label,[switch]$ThroughUI){
 $source=$fixture.Source;$hash=(Get-FileHash -LiteralPath $source).Hash
 $target=OpenTarget $label
 if($ThroughUI){
  if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$source+'")')){throw 'UI preview failed'}
  if(-not $app.Eval('UXApplySwiftWorkspace()')){throw ('UI apply failed: '+$app.Eval('Forms!frmUXWorkspaceImport!lblStatus.Caption'))}
  $result=$app.Eval('Forms!frmUXWorkspaceImport!lblStatus.Caption')
  $log=$app.Eval('Forms!frmUXWorkspaceImport!txtProgress.Value')
  $log|Set-Content (Join-Path $folder 'ui-progress.log')
  foreach($phase in @('Import started','Copying Report','records copied','Verifying all fields','records checked','Import committed successfully')){if($log -notlike ('*'+$phase+'*')){throw ('Progress missing '+$phase)}}
  if($app.Eval('UXWorkspaceImportBusy()') -or -not $app.Eval('Forms!frmUXWorkspaceImport!cmdClose.Enabled') -or $app.Eval('Forms!frmUXWorkspaceImport!cmdApply.Enabled')){throw 'Success UI state failed'}
  $result='PASS: '+$result
 }else{$result=CallImport $source}
 if($result -notlike 'PASS:*'){throw ($label+': '+$result)}
 $app.CloseCurrentDatabase()
 if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw ($label+': source modified')}
 $checked=CompareTransfer $source $target
 $app.OpenCurrentDatabase($target,$true);$reports=$app.Eval('DCount("*","Report")');$app.CloseCurrentDatabase()
 return @{Schema=$fixture.Schema;Variant=$fixture.Variant;Kind=$fixture.Kind;Definition=$fixture.Definition;Source=$source;SourceSHA256=$hash;Result=$result;Reports=$reports;RowsChecked=$checked;Target=$target}
}
try{
 foreach($fixture in $fixtures.Values){
  $result=ImportOne $fixture ('schema-'+$fixture.Variant) -ThroughUI:($fixture.Schema -eq $current)
  $results+=$result
  Write-Output ('PASS schema '+$fixture.Variant+' ('+$result.Kind+'): '+$result.RowsChecked+' source rows matched exactly, '+$result.Reports+' reports, source unchanged')
 }
 foreach($fixture in $kept){
  $result=ImportOne $fixture ('kept-'+$fixture.Schema)
  if($fixture.OriginalSHA256 -and (Get-FileHash -LiteralPath $fixture.Definition).Hash -ne $fixture.OriginalSHA256){throw ('Kept release changed: '+$fixture.Definition)}
  $keptResults+=$result
  Write-Output ('PASS kept schema '+$result.Schema+': '+$result.RowsChecked+' rows matched exactly from '+$fixture.Definition)
 }
 #Test populated-destination rejection and rollback after rows have been copied.
 $source=$fixtures[@($fixtures.Keys)[-1]].Source
 $target=OpenTarget 'guards'
 $bootstrap=$app.Eval('DLookup("CommandName","Workspace")')
 $bad=Join-Path $folder 'missing-report-source.accdb';Copy-Item -LiteralPath $source -Destination $bad
 $db=$dao.OpenDatabase($bad)
 try{
  foreach($relation in @($db.Relations|ForEach-Object {$_.Name})){if($db.Relations.Item($relation).Table -eq 'Report' -or $db.Relations.Item($relation).ForeignTable -eq 'Report'){$db.Relations.Delete($relation)}}
  $db.Execute('DROP TABLE Report',128)
 }finally{$db.Close()}
 if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$bad+'")')){throw 'Failure preview failed'}
 if($app.Eval('UXApplySwiftWorkspace()')){throw 'Broken source UI import succeeded'}
 if($app.Eval('UXWorkspaceImportBusy()') -or -not $app.Eval('Forms!frmUXWorkspaceImport!cmdClose.Enabled') -or -not $app.Eval('Forms!frmUXWorkspaceImport!cmdApply.Enabled')){throw 'Failure UI state failed'}
 if($app.Eval('Forms!frmUXWorkspaceImport!txtProgress.Value') -notlike '*Target changes rolled back*'){throw 'Failure log missing rollback'}
 $app.DoCmd.Close(2,'frmUXWorkspaceImport',2)
 $rejected=(CallImport $bad) -like 'REJECTED:*'
 if(-not $rejected -or $app.Eval('DLookup("CommandName","Workspace")') -ne $bootstrap -or $app.Eval('DCount("*","ReportRevision")') -ne 0){throw 'Failed import did not restore bootstrap/empty data'}
 if($app.Eval('TryWorkspaceImport("'+$source+'","changed-hash")') -notlike 'REJECTED:*'){throw 'Changed-source guard failed'}
 foreach($guard in @('future','unknown')){
  $invalid=Join-Path $folder ($guard+'-source.accdb');Copy-Item -LiteralPath $source -Destination $invalid
  $db=$dao.OpenDatabase($invalid)
  try{if($guard -eq 'future'){$db.Execute('UPDATE AppMetadata SET SchemaVersion=999',128)}else{$db.Execute('CREATE TABLE FutureData (Item Long)',128)}}finally{$db.Close()}
  if((CallImport $invalid) -notlike 'REJECTED:*' -or $app.Eval('DLookup("CommandName","Workspace")') -ne $bootstrap){throw ($guard+' guard failed')}
 }
 #A failure inside a post-copy migration (here a contributor whose audit identity
 #differs) must reach the import's own rollback, not an unhandled run-time dialog.
 $auditSource=if($fixtures.Contains('29')){$fixtures['29'].Source}else{$source}
 if([int](Scalar $auditSource 'SELECT Count(*) FROM AuditActor') -gt 0){
  $conflict=Join-Path $folder 'audit-identity-source.accdb';Copy-Item -LiteralPath $auditSource -Destination $conflict
  $db=$dao.OpenDatabase($conflict)
  try{$db.Execute("UPDATE AuditActor SET ProfileGuid='{00000000-0000-0000-0000-000000000001}' WHERE ActorID=(SELECT Min(ActorID) FROM AuditActor)",128)}finally{$db.Close()}
  $outcome=CallImport $conflict
  if($outcome -notlike 'REJECTED:*Contributor identity differs*' -or $app.Eval('DLookup("CommandName","Workspace")') -ne $bootstrap -or $app.Eval('DCount("*","Report")') -ne 0){throw ('Audit identity conflict was not rejected cleanly: '+$outcome)}
  $auditGuard=$true
 }
 if((CallImport $source) -notlike 'PASS:*'){throw 'Valid import after rollback failed'}
 $rows=$app.Eval('DCount("*","Report")')
 if((CallImport $source) -notlike 'REJECTED:*' -or $app.Eval('DCount("*","Report")') -ne $rows){throw 'Populated target guard failed'}
 $app.CloseCurrentDatabase()
 $receipt=@{Status='PASS';Build=$BuildPath;BuildSchema=$current;Fixtures=$results;KeptReleases=$keptResults;HistoricalFixturesRetested=(-not $CurrentFixturesOnly);Rollback=$true;ChangedSourceRejected=$true;PopulatedTargetRejected=$true;UnknownTableRejected=$true;FutureSchemaRejected=$true;UIEntryPointTested=$true;AuditIdentityConflictRejected=($auditGuard -eq $true);Evidence=$folder}
 $receipt|ConvertTo-Json -Depth 7|Set-Content (Join-Path $folder 'matrix.json')
 $receipt|ConvertTo-Json -Depth 7|Set-Content (Join-Path $root 'build/r115-data-import-verification.json')
 Write-Output ('PASS: '+$results.Count+' schema fixtures ('+(($results|ForEach-Object {$_.Schema})-join ',')+') and '+$keptResults.Count+' kept releases; production import, independent exact comparison, bootstrap rollback, changed-source, unknown-table, future-schema and populated-target guards; '+$folder)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
