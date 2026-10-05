param([Parameter(Mandatory=$true)][string]$TemplatePath,
[string]$Schema20Path,
[string]$Schema21Path)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
if(-not $Schema20Path){$Schema20Path=Join-Path $root 'output/Swift_NAVFIT_0.21.0-major-foundations.accdb'}
$folder=Join-Path $root ('build/tests/upgrade-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
$source=Join-Path $folder 'synthetic schema20.accdb';Copy-Item -LiteralPath $Schema20Path -Destination $source
$app=New-Object -ComObject Access.Application
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($source,$true)
if($app.Eval('DLookup("SchemaVersion","AppMetadata")') -ne 20){throw 'Fixture must be a real schema-20 distribution'}
$app.LoadFromText(5,'MajorHarness',(Join-Path $root 'tests/MajorHarness.bas'))
$result=$app.Eval('TestMajor()');if(-not $result.StartsWith('PASS:')){throw $result}
$app.DoCmd.DeleteObject(5,'MajorHarness');$app.CloseCurrentDatabase()
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
$sourceHash=(Get-FileHash -LiteralPath $source).Hash;$templateHash=(Get-FileHash -LiteralPath $TemplatePath).Hash
$destination=Join-Path $folder 'migrated schema21.accdb'
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'upgrade-workspace.ps1') -SourcePath $source -TemplatePath $TemplatePath -DestinationPath $destination
if($LASTEXITCODE -ne 0){throw 'Schema-20 migration failed'}
$evidence=Get-Content -LiteralPath ($destination+'.migration.json') -Raw|ConvertFrom-Json
if(-not $evidence.Reopened -or $evidence.Status -notmatch 'schema 20 to 21'){throw 'Migration evidence incomplete'}
$destHash=(Get-FileHash -LiteralPath $destination).Hash
$ErrorActionPreference='Continue'
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'upgrade-workspace.ps1') -SourcePath $source -TemplatePath $TemplatePath -DestinationPath $destination *> (Join-Path $folder 'overwrite-rejected.log')
$rejectedExit=$LASTEXITCODE;$ErrorActionPreference='Stop'
if($rejectedExit -eq 0 -or (Get-FileHash -LiteralPath $destination).Hash -ne $destHash -or (Get-Content (Join-Path $folder 'overwrite-rejected.log') -Raw) -notmatch 'Destination exists'){throw 'Destination overwrite guard failed'}
if(-not $Schema21Path){$Schema21Path=$destination}
$source21Hash=(Get-FileHash -LiteralPath $Schema21Path).Hash
$second=Join-Path $folder 'same-schema copy.accdb'
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'upgrade-workspace.ps1') -SourcePath $Schema21Path -TemplatePath $TemplatePath -DestinationPath $second
if($LASTEXITCODE -ne 0){throw 'Same-schema transfer failed'}
if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash -or (Get-FileHash -LiteralPath $TemplatePath).Hash -ne $templateHash -or (Get-FileHash -LiteralPath $destination).Hash -ne $destHash){throw 'Migration changed a source/template'}
if((Get-FileHash -LiteralPath $Schema21Path).Hash -ne $source21Hash){throw 'Schema-21 source changed'}
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
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'upgrade-workspace.ps1') -SourcePath $source -TemplatePath $Schema21Path -DestinationPath $nonblankDestination *> (Join-Path $folder 'nonblank-rejected.log')
$nonblankExit=$LASTEXITCODE;$ErrorActionPreference='Stop'
if($nonblankExit -eq 0 -or (Test-Path $nonblankDestination) -or (Get-Content (Join-Path $folder 'nonblank-rejected.log') -Raw) -notmatch 'Target must be blank' -or (Get-FileHash -LiteralPath $Schema21Path).Hash -ne $source21Hash){throw 'Populated-template preservation guard failed'}
$manifest=[ordered]@{Status='PASS';Fixture=$result;Schema20Source=$source;SourceSHA256=$sourceHash;TemplateSHA256=$templateHash;Schema20Migration=$evidence;Schema21Migration=(Get-Content ($second+'.migration.json') -Raw|ConvertFrom-Json);OverwriteRejected=$true;UnknownTableRejected=$true;PopulatedTemplateRejected=$true;Schema21SourceSHA256=$source21Hash;SourceAndTemplateUnchanged=$true}
$manifest|ConvertTo-Json -Depth 8|Set-Content (Join-Path $folder 'verification.json') -Encoding utf8
Write-Output ('PASS: populated schema 20 to 21 and 21 to 21, exact row/field comparison, reopen, overwrite/unknown-table/populated-template rejection and source/template hashes; '+$folder)
