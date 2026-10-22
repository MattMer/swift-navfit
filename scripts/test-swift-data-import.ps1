param([Parameter(Mandatory=$true)][string]$BuildPath,[ValidateRange(1,27)][int]$MinimumSchema=1,[string]$CurrentSchemaPath,[switch]$CurrentFixturesOnly)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$folder=Join-Path $root ('build/tests/data-import-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
$dao=New-Object -ComObject DAO.DBEngine.120
$fixtures=@{}
#Retained actual distributions exercise historical schema boundaries, not files
#whose version number has merely been edited to resemble an earlier schema.
$paths=@()
if($CurrentFixturesOnly){
 if(-not $CurrentSchemaPath){throw 'CurrentFixturesOnly requires -CurrentSchemaPath'}
}else{
$paths=@(Get-ChildItem (Join-Path $root 'output') -Filter *.accdb | Where-Object {$_.Name -notmatch 'candidate|rc\d|tested|tools-rc' } | Sort-Object Name)
$paths+=@(Get-ChildItem (Join-Path $root 'output/Swift_NAVFIT_0.23.0-preview'),(Join-Path $root 'output/Swift_NAVFIT_0.24.0-preview') -Filter '*DEMO.accdb')
$paths+=Get-Item (Join-Path $root 'output/ux/Swift_NAVFIT_UI_Demo_2026-10-16-r114.accdb')
$paths+=Get-Item (Join-Path $root 'output/ux/Swift_NAVFIT_UI_Demo_2026-10-06-r56.accdb'),(Join-Path $root 'output/ux/Swift_NAVFIT_UI_Demo_2026-10-07-r89.accdb')
}
if($CurrentSchemaPath){$paths+=Get-Item -LiteralPath $CurrentSchemaPath}
foreach($file in $paths){
 $db=$dao.OpenDatabase($file.FullName,$false,$true)
 try{$r=$db.OpenRecordset('SELECT SchemaVersion FROM AppMetadata',4);$v=[int]$r.Fields.Item(0).Value;$r.Close();$fixtures[$v]=$file.FullName}finally{$db.Close()}
}
$results=@();$app=New-Object -ComObject Access.Application
function OpenTarget([string]$name){
 $target=Join-Path $folder ($name+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $target
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
 foreach($form in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$form,2)}
 $app.LoadFromText(5,'WorkspaceImportHarness',(Join-Path $root 'tests/WorkspaceImportHarness.bas'))
 [void]$app.Eval('ConfigureWorkspace("SYNTHETIC IMPORT BOOTSTRAP","Tester")')
 return $target
}
function CallImport([string]$source){return $app.Eval('TryWorkspaceImport("'+$source.Replace('"','""')+'")')}
try{
 foreach($v in @($fixtures.Keys|Where-Object {$_ -ge $MinimumSchema}|Sort-Object)){
  $source=$fixtures[$v];$hash=(Get-FileHash -LiteralPath $source).Hash
  $target=OpenTarget ('schema-'+$v)
  if($v -eq 27){
   if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$source+'")')){throw 'UI preview failed'}
   if(-not $app.Eval('UXApplySwiftWorkspace()')){throw 'UI apply failed'}
   $result=$app.Eval('Forms!frmUXWorkspaceImport!lblStatus.Caption')
   $log=$app.Eval('Forms!frmUXWorkspaceImport!txtProgress.Value')
   $log|Set-Content (Join-Path $folder 'ui-progress.log')
   foreach($phase in @('Import started','Copying Report','records copied','Verifying all fields','records checked','Import committed successfully')){if($log -notlike ('*'+$phase+'*')){throw ('Progress missing '+$phase)}}
   if($app.Eval('UXWorkspaceImportBusy()') -or -not $app.Eval('Forms!frmUXWorkspaceImport!cmdClose.Enabled') -or $app.Eval('Forms!frmUXWorkspaceImport!cmdApply.Enabled')){throw 'Success UI state failed'}
  }else{$result=CallImport $source}
  if($result -notlike 'PASS:*'){throw $result}
  $count=$app.Eval('DCount("*","Report")')
  $app.CloseCurrentDatabase()
  if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw 'Source modified'}
  $results+=@{Schema=$v;Source=$source;SourceSHA256=$hash;Result=$result;Reports=$count;Target=$target}
  Write-Output ('PASS schema '+$v+': '+$count+' reports, exact field verification and unchanged source')
 }
 #Test populated-destination rejection and rollback after rows have been copied.
 $source=if($CurrentFixturesOnly){$fixtures[@($fixtures.Keys)[0]]}else{$fixtures[25]}
 $target=OpenTarget 'guards'
 $bootstrap=$app.Eval('DLookup("CommandName","Workspace")')
 $bad=Join-Path $folder 'missing-report-source.accdb';Copy-Item -LiteralPath $source -Destination $bad
 $db=$dao.OpenDatabase($bad)
 try{
  foreach($relation in @($db.Relations|ForEach-Object {$_.Name})){if($db.Relations.Item($relation).Table -eq 'Report' -or $db.Relations.Item($relation).ForeignTable -eq 'Report'){$db.Relations.Delete($relation)}}
  $db.Execute('DROP TABLE Report',128)
 }finally{$db.Close()}
 $rejected=$false
 if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$bad+'")')){throw 'Failure preview failed'}
 if($app.Eval('UXApplySwiftWorkspace()')){throw 'Broken source UI import succeeded'}
 if($app.Eval('UXWorkspaceImportBusy()') -or -not $app.Eval('Forms!frmUXWorkspaceImport!cmdClose.Enabled') -or -not $app.Eval('Forms!frmUXWorkspaceImport!cmdApply.Enabled')){throw 'Failure UI state failed'}
 if($app.Eval('Forms!frmUXWorkspaceImport!txtProgress.Value') -notlike '*Target changes rolled back*'){throw 'Failure log missing rollback'}
 $app.DoCmd.Close(2,'frmUXWorkspaceImport',2)
 $rejected=(CallImport $bad) -like 'REJECTED:*'
 if(-not $rejected -or $app.Eval('DLookup("CommandName","Workspace")') -ne $bootstrap -or $app.Eval('DCount("*","ReportRevision")') -ne 0){throw 'Failed import did not restore bootstrap/empty data'}
 $stale=$false
 $stale=$app.Eval('TryWorkspaceImport("'+$source+'","changed-hash")') -like 'REJECTED:*'
 if(-not $stale){throw 'Changed-source guard failed'}
 foreach($guard in @('future','unknown')){
  $invalid=Join-Path $folder ($guard+'-source.accdb');Copy-Item -LiteralPath $source -Destination $invalid
  $db=$dao.OpenDatabase($invalid)
  try{if($guard -eq 'future'){$db.Execute('UPDATE AppMetadata SET SchemaVersion=999',128)}else{$db.Execute('CREATE TABLE FutureData (Item Long)',128)}}finally{$db.Close()}
  if((CallImport $invalid) -notlike 'REJECTED:*' -or $app.Eval('DLookup("CommandName","Workspace")') -ne $bootstrap){throw ($guard+' guard failed')}
 }
 if((CallImport $source) -notlike 'PASS:*'){throw 'Valid import after rollback failed'}
 $rows=$app.Eval('DCount("*","Report")');$rejected=$false
 $rejected=(CallImport $source) -like 'REJECTED:*'
 if(-not $rejected -or $app.Eval('DCount("*","Report")') -ne $rows){throw 'Populated target guard failed'}
 $app.CloseCurrentDatabase()
 $results|ConvertTo-Json -Depth 6|Set-Content (Join-Path $folder 'matrix.json')
 @{Status='PASS';Fixtures=$results;HistoricalFixturesRetested=(-not $CurrentFixturesOnly);Rollback=$true;ChangedSourceRejected=$true;PopulatedTargetRejected=$true;UnknownTableRejected=$true;FutureSchemaRejected=$true;UIEntryPointTested=$true}|ConvertTo-Json -Depth 7|Set-Content (Join-Path $root 'build/r115-data-import-verification.json')
 Write-Output ('PASS: '+$results.Count+' schema fixtures; historical retest: '+(-not $CurrentFixturesOnly)+'; production import, bootstrap rollback, changed-source and populated-target guards; '+$folder)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
