param([Parameter(Mandatory=$true)][string]$TemplatePath,
[string]$Schema20Path,
[Alias("Schema21Path")][string]$SourceWorkspacePath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$folder=Join-Path $root ('build/tests/upgrade-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
$source=Join-Path $folder 'synthetic schema20.accdb'
#The populated schema-20 source is generated from tests/historical-schemas by the
#import matrix's own builder; old release files are not needed (AGENTS.md).
if($Schema20Path){Copy-Item -LiteralPath $Schema20Path -Destination $source}
else{
 $made=& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'test-swift-data-import.ps1') -BuildPath $TemplatePath -Schemas 20 -FixturePath $source 2>&1
 if($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $source)){throw ('Schema-20 fixture could not be generated: '+($made -join ' '))}
}
$checkDAO=New-Object -ComObject DAO.DBEngine.120;$checkDB=$checkDAO.OpenDatabase($source,$false,$true)
try{$r=$checkDB.OpenRecordset('SELECT SchemaVersion FROM AppMetadata',4);$v=[int]$r.Fields.Item(0).Value;$r.Close();$reports=[int]$checkDB.OpenRecordset('SELECT Count(*) FROM Report',4).Fields.Item(0).Value}finally{$checkDB.Close()}
if($v -ne 20 -or $reports -lt 1){throw 'Fixture must be a populated schema-20 workspace'}
$result='populated synthetic schema-20 source with '+$reports+' reports'
$migrationDAO=New-Object -ComObject DAO.DBEngine.120;$migrationDB=$migrationDAO.OpenDatabase([IO.Path]::GetFullPath($TemplatePath),$false,$true)
try{$migrationRows=$migrationDB.OpenRecordset('SELECT SchemaVersion FROM AppMetadata',4);$targetSchema=[int]$migrationRows.Fields.Item(0).Value;$migrationRows.Close()}finally{$migrationDB.Close()}
$sourceHash=(Get-FileHash -LiteralPath $source).Hash;$templateHash=(Get-FileHash -LiteralPath $TemplatePath).Hash
$destination=Join-Path $folder 'migrated current.accdb'
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'upgrade-workspace.ps1') -SourcePath $source -TemplatePath $TemplatePath -DestinationPath $destination
if($LASTEXITCODE -ne 0){throw 'Schema-20 migration failed'}
$evidence=Get-Content -LiteralPath ($destination+'.migration.json') -Raw|ConvertFrom-Json
if(-not $evidence.Reopened -or $evidence.Status -notmatch ('schema 20 to '+$targetSchema)){throw 'Migration evidence incomplete'}
$destHash=(Get-FileHash -LiteralPath $destination).Hash
$ErrorActionPreference='Continue'
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'upgrade-workspace.ps1') -SourcePath $source -TemplatePath $TemplatePath -DestinationPath $destination *> (Join-Path $folder 'overwrite-rejected.log')
$rejectedExit=$LASTEXITCODE;$ErrorActionPreference='Stop'
if($rejectedExit -eq 0 -or (Get-FileHash -LiteralPath $destination).Hash -ne $destHash -or (Get-Content (Join-Path $folder 'overwrite-rejected.log') -Raw) -notmatch 'Destination exists'){throw 'Destination overwrite guard failed'}
if(-not $SourceWorkspacePath){$SourceWorkspacePath=$destination}
$source21Hash=(Get-FileHash -LiteralPath $SourceWorkspacePath).Hash
$second=Join-Path $folder 'same-schema copy.accdb'
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'upgrade-workspace.ps1') -SourcePath $SourceWorkspacePath -TemplatePath $TemplatePath -DestinationPath $second
if($LASTEXITCODE -ne 0){throw 'Same-schema transfer failed'}
if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash -or (Get-FileHash -LiteralPath $TemplatePath).Hash -ne $templateHash -or (Get-FileHash -LiteralPath $destination).Hash -ne $destHash){throw 'Migration changed a source/template'}
if((Get-FileHash -LiteralPath $SourceWorkspacePath).Hash -ne $source21Hash){throw 'Schema-21 source changed'}
$badSource=Join-Path $folder 'unknown-table source.accdb';Copy-Item -LiteralPath $source -Destination $badSource
$dao=New-Object -ComObject DAO.DBEngine.120;$db=$dao.OpenDatabase($badSource)
try{$db.Execute('CREATE TABLE UnrecognizedSynthetic (Item Long)',128)}finally{$db.Close()}
$badHash=(Get-FileHash -LiteralPath $badSource).Hash
$badDestination=Join-Path $folder 'must-not-exist.accdb'
$ErrorActionPreference='Continue'
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'upgrade-workspace.ps1') -SourcePath $badSource -TemplatePath $TemplatePath -DestinationPath $badDestination *> (Join-Path $folder 'unknown-table-rejected.log')
$badExit=$LASTEXITCODE;$ErrorActionPreference='Stop'
if($badExit -eq 0 -or (Test-Path $badDestination) -or (Get-Content (Join-Path $folder 'unknown-table-rejected.log') -Raw) -notmatch 'Unrecognized source table' -or (Get-FileHash -LiteralPath $badSource).Hash -ne $badHash){throw 'Unknown-source-table preservation guard failed'}
$nonblankDestination=Join-Path $folder 'nonblank-must-not-exist.accdb'
$ErrorActionPreference='Continue'
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'upgrade-workspace.ps1') -SourcePath $source -TemplatePath $SourceWorkspacePath -DestinationPath $nonblankDestination *> (Join-Path $folder 'nonblank-rejected.log')
$nonblankExit=$LASTEXITCODE;$ErrorActionPreference='Stop'
if($nonblankExit -eq 0 -or (Test-Path $nonblankDestination) -or (Get-Content (Join-Path $folder 'nonblank-rejected.log') -Raw) -notmatch 'Target must be blank' -or (Get-FileHash -LiteralPath $SourceWorkspacePath).Hash -ne $source21Hash){throw 'Populated-template preservation guard failed'}
$manifest=[ordered]@{Status='PASS';Fixture=$result;Schema20Source=$source;SourceSHA256=$sourceHash;TemplateSHA256=$templateHash;Schema20Migration=$evidence;Schema21Migration=(Get-Content ($second+'.migration.json') -Raw|ConvertFrom-Json);OverwriteRejected=$true;UnknownTableRejected=$true;PopulatedTemplateRejected=$true;Schema21SourceSHA256=$source21Hash;SourceAndTemplateUnchanged=$true}
$manifest|ConvertTo-Json -Depth 8|Set-Content (Join-Path $folder 'verification.json') -Encoding utf8
Write-Output ('PASS: populated schema 20 to '+$targetSchema+' and supplied workspace to '+$targetSchema+', exact row/field comparison, reopen, overwrite/unknown-table/populated-template rejection and source/template hashes; '+$folder)
