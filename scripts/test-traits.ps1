param([Parameter(Mandatory=$true)][string]$BuildPath,[string]$LegacyPath,[switch]$LoadCurrentSource)
# Traits suite: persistence, rollback, legacy restore and fresh-session recovery; exact
# catalogue wording and production-migration transfer; grid events, queue, summary and retry.
# -LegacyPath adds the schema-30 catalogue transfer case; -LoadCurrentSource loads the
# current modUXTraitGrid into the retry case's copy.
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'suite-common.ps1')
. (Join-Path $PSScriptRoot 'test-contributor-context.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash

Invoke-SuiteCase 'traits' {
 $scratch=Join-Path $root ('build/tests/traits-'+[guid]::NewGuid()+'.accdb')
 Copy-Item -LiteralPath $BuildPath -Destination $scratch
 $app=$null
 try {
  $app=New-Object -ComObject Access.Application
  $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
  $app.LoadFromText(5,'TraitsHarness',(Join-Path $root 'tests/TraitsHarness.bas'))
  $app.DoCmd.RunCommand(126)
  $result=$app.Eval('TestTraits()')
  if($result -notlike 'PASS:*'){throw $result}
  $app.CloseCurrentDatabase();$app.Quit(2)
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
  $app=New-Object -ComObject Access.Application
  $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
  SelectSyntheticContributor $app
  if(-not $app.Eval('VerifyTraitRecovery()')){throw 'Trait recovery in new session failed'}
  $app.CloseCurrentDatabase()
  @{Status=$result;RecoveryInNewSession=$true;BuildSHA256=$buildHash;TestDatabase=$scratch;Scope='Trait persistence, rollback, legacy restore and recovery'} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/traits-results.json')
  Write-Output $result
  Write-Output 'PASS: invalid pending trait recovery in a new Access session'
 } finally {
  if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 }
}

# Exact catalogue wording, runtime rows and constraints, then production-migration
# transfer of current, conflicting, malformed and (with -LegacyPath) schema-30 sources.
Invoke-SuiteCase 'trait-catalogue' {
 $hash=(Get-FileHash -LiteralPath $BuildPath).Hash
 $scratch=Join-Path $root ('build/tests/catalogue-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
 $catalogue=Get-Content (Join-Path $root 'src/trait-catalogue.json') -Raw|ConvertFrom-Json
 $app=$null
 try{
  $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
  if($app.Eval('TraitCatalogueHash()') -cne $catalogue.SourceSHA256){throw 'Embedded catalogue identity changed'}
  $checks=1
  $db=$app.CurrentDb();$row=$db.OpenRecordset('SELECT Count(*) FROM TraitDefinition',4)
  if($row.Fields.Item(0).Value -ne 147){throw 'Expected 147 versioned runtime trait rows'};$checks++;$row.Close()
  foreach($trait in $catalogue.Traits){
   $prefix='("'+$trait.FormKind+'",'+$trait.Block
   if($app.Eval('CatalogueTraitTitle'+$prefix+')') -cne $trait.Title){throw 'Trait title changed'};$checks++
   if($app.Eval('CatalogueTraitDescription'+$prefix+',"")') -cne [string]$trait.GeneralDescription){throw 'Trait general description changed'};$checks++
   foreach($mark in @('NOB','1.0','2.0','3.0','4.0','5.0')){
    if($app.Eval('CatalogueTraitDescription'+$prefix+',"'+$mark+'")') -cne [string]$trait.Marks.$mark){throw 'Trait mark description changed'};$checks++
   }
   foreach($mark in @('GENERAL','NOB','1.0','2.0','3.0','4.0','5.0')){
    $row=$db.OpenRecordset("SELECT * FROM TraitDefinition WHERE FormKind='"+$trait.FormKind+"' AND BlockNumber="+$trait.Block+" AND Mark='"+$mark+"' AND CatalogueVersion=1",4)
    if($row.EOF){throw 'Missing runtime row'}
    $description=if($mark -eq 'GENERAL'){$trait.GeneralDescription}else{$trait.Marks.$mark}
    foreach($pair in @(@('TraitTitle',$trait.Title),@('DescriptionText',$description),@('SourceName',$catalogue.Source),@('SourceSHA256',$catalogue.SourceSHA256),@('SourceTraitID',$trait.SourceTraitID))){
     $actual=$row.Fields.Item($pair[0]).Value
     if($null -eq $pair[1]){if($actual -isnot [DBNull]){throw 'Reference Null changed'}}elseif([string]$actual -cne [string]$pair[1]){throw ('Reference field differs: '+$pair[0])};$checks++
    };$row.Close()
   }
  }
  # Constraints must reject duplicate identities and invalid form/block/mark data.
  foreach($sql in @("INSERT INTO TraitDefinition SELECT * FROM TraitDefinition WHERE FormKind='EVAL' AND BlockNumber=33 AND Mark='GENERAL'", "UPDATE TraitDefinition SET Mark='INVALID' WHERE FormKind='EVAL' AND BlockNumber=33 AND Mark='GENERAL'", "UPDATE TraitDefinition SET BlockNumber=32 WHERE FormKind='EVAL' AND BlockNumber=33 AND Mark='GENERAL'")){
   $rejected=$false;try{$db.Execute($sql,128)}catch{$rejected=$true};if(-not $rejected){throw 'Trait identity/domain constraint failed'};$checks++
  }
  $db.Execute("DELETE FROM TraitDefinition WHERE FormKind='EVAL' AND BlockNumber=33 AND Mark='GENERAL'",128)
  $app.LoadFromText(5,'TraitsHarness',(Join-Path $root 'tests/TraitsHarness.bas'))
  $rejected=[bool]$app.Run('TraitDefinitionMissingRejected')
  if(-not $rejected){throw 'Missing required definition was silently ignored'};$checks++
  $db.Close()
  Write-Output ('PASS: '+$checks+' exact embedded trait-catalogue fidelity checks')
 }finally{if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
 if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash){throw 'Catalogue test changed source artifact'}
 
 # Stateful reference preservation uses the production migration engine, not SQL
 # copy helpers. Each controller releases its file before the next phase opens it.
 $folder=Join-Path $root ('build/tests/trait-transfer-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
 $source=Join-Path $folder 'source.accdb';Copy-Item -LiteralPath $BuildPath -Destination $source
 $app=$null
 try{
  Write-Output ('Controller: trait transfer runner; database: '+$source)
  $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($source,$true)
  [void]$app.Eval('ConfigureWorkspace("SYNTHETIC TRAIT TRANSFER","Synthetic Contributor")')
  [void]$app.Eval('CommitDraft(0,"EVAL","Synthetic","Synthetic Member",Null,Null,"","Synthetic transfer","Synthetic narrative")')
  $db=$app.CurrentDb()
  $db.Execute("INSERT INTO TraitDefinition (FormKind,BlockNumber,Mark,CatalogueVersion,TraitTitle,DescriptionText,SourceName,SourceSHA256,SourceTraitID) SELECT FormKind,BlockNumber,Mark,2,TraitTitle,DescriptionText,SourceName,SourceSHA256,SourceTraitID FROM TraitDefinition WHERE FormKind='EVAL' AND BlockNumber=33 AND Mark='GENERAL'",128)
  $db.Close()
 }finally{if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null}}
 $sourceHash=(Get-FileHash -LiteralPath $source).Hash
 $conflict=Join-Path $folder 'conflict.accdb';Copy-Item -LiteralPath $source -Destination $conflict
 $dao=New-Object -ComObject DAO.DBEngine.120
 try{$db=$dao.OpenDatabase($conflict);$db.Execute("UPDATE TraitDefinition SET TraitTitle='Synthetic conflict' WHERE FormKind='EVAL' AND BlockNumber=33 AND Mark='GENERAL' AND CatalogueVersion=1",128);$db.Close()}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($dao)}
 $cases=@(@('current',$source,148,$true),@('conflict',$conflict,147,$false))
 $malformed=Join-Path $folder 'malformed-columns.accdb';Copy-Item -LiteralPath $source -Destination $malformed
 $dao=New-Object -ComObject DAO.DBEngine.120
 try{
  $db=$dao.OpenDatabase($malformed)
  $db.TableDefs.Item('TraitDefinition').ValidationRule=''
  $db.Execute('ALTER TABLE TraitDefinition DROP CONSTRAINT PK_TraitDefinition',128)
  $db.Execute('ALTER TABLE TraitDefinition ALTER COLUMN BlockNumber TEXT(32)',128)
  $db.Close()
 }finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($dao)}
 $cases+=,@('malformed',$malformed,147,$false)
 $legacyHash=$null
 if($LegacyPath){
  $legacyHash=(Get-FileHash -LiteralPath $LegacyPath).Hash
  $legacy=Join-Path $folder 'legacy-source.accdb';Copy-Item -LiteralPath $LegacyPath -Destination $legacy
  try{
   Write-Output ('Controller: trait legacy source setup; database: '+$legacy)
   $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($legacy,$true)
   if($app.Eval('DLookup("SchemaVersion","AppMetadata")') -ne 30){throw 'Use an actual schema-30 blank for the legacy proof'}
   [void]$app.Eval('ConfigureWorkspace("SYNTHETIC TRAIT LEGACY","Synthetic Contributor")')
   [void]$app.Eval('CommitDraft(0,"CHIEFEVAL","Synthetic","Synthetic Member",Null,Null,"","Synthetic legacy","Synthetic legacy narrative")')
  }finally{if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null}}
  $cases+=,@('legacy30',$legacy,147,$true)
 }
 $results=@()
 foreach($case in $cases){
  $target=Join-Path $folder ($case[0]+'-target.accdb');Copy-Item -LiteralPath $BuildPath -Destination $target
  $inputHash=(Get-FileHash -LiteralPath $case[1]).Hash
  try{
   Write-Output ('Controller: production trait transfer; database: '+$target)
   $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
   $result=[string]$app.Eval('MigrateWorkspace("'+([string]$case[1]).Replace('"','""')+'",False,False,"'+$inputHash+'")')
   if(($result -like 'PASS:*') -ne $case[3]){throw ('Unexpected migration result: '+$case[0]+'; '+$result)}
   if($app.Eval('DCount("*","TraitDefinition")') -ne $case[2]){throw 'Reference versions lost or conflicting data retained'}
   if($case[3]){
    if($app.Eval('DCount("*","Report")') -ne 1 -or $app.Eval('DCount("*","ReportRevision")') -ne 1){throw 'Populated report/revision transfer failed'}
    if($app.Eval('CatalogueTraitTitle("EVAL",33)') -cne ($catalogue.Traits|Where-Object {$_.FormKind -eq 'EVAL' -and $_.Block -eq 33}).Title){throw 'Import changed current official wording'}
   }else{
    $expectedFailure=if($case[0] -eq 'malformed'){'*Invalid trait reference column types*'}else{'*Conflicting trait reference data*'}
    if($app.Eval('DCount("*","Report")') -ne 0 -or $app.Eval('DCount("*","Workspace")') -ne 0 -or $result -notlike $expectedFailure){throw 'Invalid source did not roll back reference/domain data'}
   }
   $results+=[pscustomobject]@{Case=$case[0];Result=$result;ExpectedReferenceRows=$case[2]}
  }finally{if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null}}
  if((Get-FileHash -LiteralPath $case[1]).Hash -ne $inputHash){throw 'Production migration changed source bytes'}
  Write-Output ('PASS: '+$case[0]+' reference preservation/domain transfer/rollback and source hash')
 }
 if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $hash -or (Get-FileHash -LiteralPath $source).Hash -ne $sourceHash){throw 'Trait verification changed protected input'}
 if($LegacyPath -and (Get-FileHash -LiteralPath $LegacyPath).Hash -ne $legacyHash){throw 'Legacy template changed'}
 @{Status='PASS';ExactChecks=$checks;BuildSHA256=$hash;LegacySHA256=$legacyHash;CheckerSHA256=(Get-FileHash -LiteralPath $PSCommandPath).Hash;HarnessSHA256=(Get-FileHash -LiteralPath (Join-Path $root 'tests/TraitsHarness.bas')).Hash;CatalogueSHA256=(Get-FileHash -LiteralPath (Join-Path $root 'src/trait-catalogue.json')).Hash;Cases=$results}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $folder 'verification.json')
 Write-Output ('Receipt: '+(Join-Path $folder 'verification.json'))
}

Invoke-HarnessCase 'trait-events' $BuildPath 'TraitEventHarness' 'TestTraitEvents()' -Exclusive $false
Invoke-HarnessCase 'trait-queue' $BuildPath 'TraitQueueHarness' 'TestTraitQueue()' -Exclusive $false -After {param($app) if($app.CurrentProject.AllForms.Item('frmUXComparison').IsLoaded){[void]$app.Eval('UXGridDiscardPending()')}}
Invoke-HarnessCase 'trait-summary' $BuildPath 'TraitSummaryHarness' 'TestTraitSummary()'
Invoke-HarnessCase 'trait-retry' $BuildPath 'UXTraitRetryHarness' 'TestTraitRetry()' -Before {param($app) foreach($opened in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$opened,2)}} -SourceModules @($(if($LoadCurrentSource){'modUXTraitGrid'}))

Complete-SuiteRun 'traits' $BuildPath $buildHash
