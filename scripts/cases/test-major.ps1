param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource,[string]$LegacyPath,[string]$OlderPath,[switch]$ChiefComments)
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$folder=Join-Path $root ('build/tests/major-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
$hashes=@{};$results=@();$app=$null;$primary=$null;$cleanup=@()
foreach($path in @($BuildPath,$LegacyPath,$OlderPath)|Where-Object {$_}){$hashes[[IO.Path]::GetFullPath($path)]=(Get-FileHash -LiteralPath $path).Hash}
$schema=Get-Content (Join-Path $root 'src/major-schema.json') -Raw|ConvertFrom-Json
$currentSchema=[int]$schema.SchemaVersion
$version=[regex]::Match((Get-Content (Join-Path $root 'src/modSwift.bas') -Raw),'APP_VERSION As String = "([^"]+)"').Groups[1].Value
$nativeFixture=Join-Path $folder 'synthetic-native.accdb'
if($ChiefComments){& (Join-Path (Split-Path $PSScriptRoot -Parent) 'make-navfit-full-fixture.ps1') -DestinationPath $nativeFixture}
function CloseMajorSession {
 if($script:app){
  try{$script:app.CloseCurrentDatabase()}catch{$script:cleanup+=$_.Exception.Message}
  try{$script:app.Quit(2)}catch{$script:cleanup+=$_.Exception.Message}
  try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:app)}catch{$script:cleanup+=$_.Exception.Message}
  $script:app=$null
 }
 foreach($name in @('db','w','replacement','newField','rel','query','oldField')){
  $variable=Get-Variable -Name $name -Scope Script -ErrorAction SilentlyContinue
  if($variable -and $variable.Value -and [Runtime.InteropServices.Marshal]::IsComObject($variable.Value)){
   try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($variable.Value)}catch{$script:cleanup+=$_.Exception.Message}
   Set-Variable -Name $name -Scope Script -Value $null
  }
 }
 [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
 if($script:cleanup.Count){throw ('Owned session cleanup: '+($script:cleanup -join '; '))}
}
function OpenMajorSession([string]$path){
 Write-Output ('Controller: major workflow/transfer runner; disposable database: '+$path)
 $script:app=New-Object -ComObject Access.Application
 $script:app.AutomationSecurity=1;$script:app.OpenCurrentDatabase($path,$true)
 foreach($module in @('MajorHarness','WorkspaceImportHarness')){$script:app.LoadFromText(5,$module,(Join-Path $root ('tests/'+$module+'.bas')))}
 $script:app.DoCmd.RunCommand(126)
}
function RelationSignatures($db){
 @($db.Relations|Where-Object {$_.Table -notlike 'MSys*' -and $_.ForeignTable -notlike 'MSys*'}|ForEach-Object {
  $rel=$_;$fields=@($rel.Fields|ForEach-Object {$_.Name+'='+$_.ForeignName})
  $rel.Table+'|'+$rel.ForeignTable+'|'+$rel.Attributes+'|'+($fields -join ',')
 }|Sort-Object)
}
function AssertMajorSchema {
 $db=$script:app.CurrentDb()
 try{
  $meta=$db.OpenRecordset('SELECT SchemaVersion FROM AppMetadata',4)
  if($meta.Fields.Item(0).Value -ne $currentSchema){throw 'Use a current-schema blank for the target; legacy templates belong in -LegacyPath/-OlderPath'}
  $meta.Close()
  $checked=0
  foreach($sql in $schema.SQL){
   if($sql -notmatch '^ALTER TABLE (\w+) ADD CONSTRAINT (FK_\w+) FOREIGN KEY \((\w+)\) REFERENCES (\w+) \((\w+)\)$'){continue}
   $rel=$db.Relations.Item($Matches[2])
   if($rel.ForeignTable -ne $Matches[1] -or $rel.Table -ne $Matches[4] -or $rel.Fields.Count -ne 1 -or $rel.Fields.Item(0).ForeignName -ne $Matches[3] -or $rel.Fields.Item(0).Name -ne $Matches[5] -or ($rel.Attributes -band 2)){throw ('Relationship differs: '+$Matches[2])}
   $checked++
  }
  if(@($db.Relations|Where-Object {$_.Name -match '^FK_Major\d+$'}).Count){throw 'Opaque relationships remain'}
  foreach($query in $schema.Queries){
   if([regex]::Replace($db.QueryDefs.Item($query.Name).SQL,'[\[\];\s]','') -ine [regex]::Replace($query.SQL,'[\[\];\s]','')){throw ('Alias SQL differs: '+$query.Name)}
  }
  Write-Output ('PASS: '+$checked+' named single-column relationship contracts and both alias definitions')
 }finally{$db.Close()}
}
function CheckMajorResult([string]$expression){
 $result=[string]$script:app.Eval($expression);Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $script:results+=@{Expression=$expression;Result=$result}
}
try{
 $candidate=Join-Path $folder 'candidate.accdb';Copy-Item -LiteralPath $BuildPath -Destination $candidate
 OpenMajorSession $candidate
 if($InstallSource){
  $db=$app.CurrentDb();$before=RelationSignatures $db
  $keys=@((Get-Content (Join-Path $root 'src/field-definitions.json') -Raw|ConvertFrom-Json).fields.key)
  $present=@($db.TableDefs.Item('Report').Fields|ForEach-Object {$_.Name})
  if(@($keys|Where-Object {$_ -notin $present}).Count){throw 'Major source installation cannot add report fields; use the current fresh blank or the focused PDF schema preparation'}
  # Test-copy preparation only. Production migration uses the fresh schema and
  # reads the older database without renaming or mutating its relationships.
  if($app.Eval('DCount("*","Report")') -ne 0){throw 'Source-install preparation requires a blank disposable copy'}
  $w=$app.DBEngine.Workspaces.Item(0);$w.BeginTrans()
  try{foreach($sql in $schema.SQL){
   if($sql -notmatch '^ALTER TABLE (\w+) ADD CONSTRAINT (FK_\w+) FOREIGN KEY \((\w+)\) REFERENCES (\w+) \((\w+)\)$'){continue}
   $name=$Matches[2];$child=$Matches[1];$field=$Matches[3]
   foreach($rel in @($db.Relations|Where-Object {$_.Name -match '^FK_Major\d+$' -and $_.ForeignTable -eq $child -and $_.Fields.Count -eq 1 -and $_.Fields.Item(0).ForeignName -eq $field})){
    $replacement=$db.CreateRelation($name,$rel.Table,$rel.ForeignTable,$rel.Attributes)
    foreach($oldField in $rel.Fields){$newField=$replacement.CreateField($oldField.Name);$newField.ForeignName=$oldField.ForeignName;$replacement.Fields.Append($newField)}
    $db.Relations.Delete($rel.Name);$db.Relations.Append($replacement)
   }
  };$w.CommitTrans()}catch{$w.Rollback();throw}
  $db.Relations.Refresh()
  if(Compare-Object $before (RelationSignatures $db)){throw 'Relationship semantics changed during test-copy preparation'}
  foreach($definition in $schema.Queries){
   if(@($db.QueryDefs|ForEach-Object {$_.Name}) -contains $definition.Name){$db.QueryDefs.Delete($definition.Name)}
   $query=$db.CreateQueryDef($definition.Name,$definition.SQL);$query.Close()
  }
  $db.Execute("UPDATE AppMetadata SET SchemaVersion=$currentSchema,AppVersion='$version'",128);$db.Close()
  foreach($module in @('modSwift','modMigration','modOfficialBatch','modMeetings','modMajorUI','modUXTraitGrid','modGroupPolicy')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}
  $app.DoCmd.RunCommand(126)
 }
 AssertMajorSchema
 CloseMajorSession
 $workflow=Join-Path $folder 'workflow.accdb';Copy-Item -LiteralPath $candidate -Destination $workflow
 OpenMajorSession $workflow
 CheckMajorResult 'TestMajor()'
 CheckMajorResult 'TestMajorRevisionAliases(True)'
 CheckMajorResult 'TestMajorAliasViews()'
 CloseMajorSession
 $templates=@($candidate)+@($LegacyPath,$OlderPath|Where-Object {$_})
 $index=0
 foreach($template in $templates){
  $index++;$source=Join-Path $folder ('source-'+$index+'.accdb');Copy-Item -LiteralPath $template -Destination $source
  OpenMajorSession $source
  $version=[int]$app.Eval('DLookup("SchemaVersion","AppMetadata")')
  CheckMajorResult 'TestMajor()'
  $oldExport=Join-Path $folder ('seed-export-'+$index+'.accdb')
  if($ChiefComments){CheckMajorResult ('SeedChiefTransfer("'+$nativeFixture+'","'+$oldExport+'","'+(Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb')+'")')}
  $db=$app.CurrentDb();$sourceRelations=RelationSignatures $db;$db.Close()
  CloseMajorSession
  $sourceHash=(Get-FileHash -LiteralPath $source).Hash
  $target=Join-Path $folder ('target-'+$version+'-'+$index+'.accdb');Copy-Item -LiteralPath $candidate -Destination $target
  OpenMajorSession $target
  [void]$app.Eval('ConfigureWorkspace("SYNTHETIC TARGET SETUP","Synthetic Tester")')
  if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$source+'")')){throw 'Settings import preview failed'}
  $db=$app.CurrentDb();$db.TableDefs.Item('MeetingProposal').Fields.Item('ProposalState').ValidationRule="='Reject imported proposals'";$db.Close()
  $rejected=[string]$app.Eval('TryWorkspaceImport("'+$source+'")')
  if($rejected -notlike 'REJECTED:*' -or $app.Eval('DCount("*","Report")') -ne 0 -or $app.Eval('DCount("*","ChangeItem")') -ne 0 -or $app.Eval('DLookup("CommandName","Workspace")') -ne 'SYNTHETIC TARGET SETUP'){throw ('Late rollback failed: '+$rejected)}
  $db=$app.CurrentDb();$db.TableDefs.Item('MeetingProposal').Fields.Item('ProposalState').ValidationRule='';$db.Close()
  if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$source+'")') -or -not $app.Eval('UXApplySwiftWorkspace()')){throw 'Settings import apply failed'}
  AssertMajorSchema
  $db=$app.CurrentDb();if(Compare-Object $sourceRelations (RelationSignatures $db)){throw 'Imported relationship semantics differ'};$db.Close()
  CheckMajorResult 'TestMajorRevisionAliases()'
  if($ChiefComments){CheckMajorResult ('VerifyChiefTransfer('+ $version +',"'+$nativeFixture+'","'+$oldExport+'","'+(Join-Path $folder ('target-export-'+$index+'.accdb'))+'","'+(Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb')+'")')}
  $repeat=[string]$app.Eval('MigrateWorkspace("'+$source+'")')
  if($repeat -notlike '*Target must be blank*'){throw ('Populated target was not rejected: '+$repeat)}
  $app.CloseCurrentDatabase();$app.OpenCurrentDatabase($target,$true)
  CheckMajorResult 'TestMajorRevisionAliases()'
  CloseMajorSession
  if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash){throw 'Import changed source'}
  $results+=@{SourceSchema=$version;SourceHash=$sourceHash;Target=$target;LateRollback=$true;Reopened=$true;AllOriginalFieldsVerified=$true}
  Write-Output ('PASS: populated schema '+$version+' -> '+$currentSchema+' Settings import, exact original-field verification, relationship equivalence, late rollback and reopen')
 }
 foreach($case in @('future','unknown-table')){
  $source=Join-Path $folder ($case+'.accdb');Copy-Item -LiteralPath $candidate -Destination $source
  $dao=New-Object -ComObject DAO.DBEngine.120
  try{$db=$dao.OpenDatabase($source);if($case -eq 'future'){$db.Execute(('UPDATE AppMetadata SET SchemaVersion='+($currentSchema+1)),128)}else{$db.Execute('CREATE TABLE UnexpectedUserData (ID LONG PRIMARY KEY)',128)};$db.Close()}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($dao)}
  $sourceHash=(Get-FileHash -LiteralPath $source).Hash
  $target=Join-Path $folder ($case+'-target.accdb');Copy-Item -LiteralPath $candidate -Destination $target
  OpenMajorSession $target
  $rejected=[string]$app.Eval('MigrateWorkspace("'+$source+'")')
  $expected=if($case -eq 'future'){'Unsupported source schema'}else{'Unrecognized source table'}
  if($rejected -notlike ('*'+$expected+'*') -or $app.Eval('DCount("*","Report")') -ne 0){throw ('Rejection failed: '+$rejected)}
  CloseMajorSession
  if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash){throw 'Rejected source changed'}
  $results+=@{Case=$case;Rejected=$true;SourceUnchanged=$true}
 }
}catch{$primary=$_}finally{
 try{CloseMajorSession}catch{if(-not $primary){$primary=$_}}
 foreach($path in $hashes.Keys){if((Get-FileHash -LiteralPath $path).Hash -ne $hashes[$path]){$cleanup+='Protected input changed: '+$path}}
}
if($primary){if($cleanup){Write-Warning ($cleanup -join '; ')};throw $primary}
if($cleanup){throw ($cleanup -join '; ')}
@{Status='PASS';InstalledSource=[bool]$InstallSource;InputHashes=$hashes;CheckerSHA256=(Get-FileHash -LiteralPath $PSCommandPath).Hash;HarnessSHA256=(Get-FileHash -LiteralPath (Join-Path $root 'tests/MajorHarness.bas')).Hash;Results=$results;Directory=$folder}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $folder 'verification.json')
Write-Output ('Receipt: '+(Join-Path $folder 'verification.json'))
