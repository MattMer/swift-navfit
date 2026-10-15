param(
 [string]$Executable='D:\Development\Swift NAVFIT\NAVFIT98A\NAVFIT98A.exe',
 [Parameter(Mandatory=$true)][string]$ReaEntryPath,
 [Parameter(Mandatory=$true)][string]$NodePath,
 [Parameter(Mandatory=$true)][string]$OutputDirectory
)
# Read-only REA CLI inspection; no package acquisition or MCP/configuration changes.
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$Executable=[IO.Path]::GetFullPath($Executable);$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $OutputDirectory){throw 'Refusing existing analysis directory'}
[void](New-Item -ItemType Directory -Path $OutputDirectory)
$before=(Get-FileHash -LiteralPath $Executable).Hash
& $NodePath $ReaEntryPath inspect-managed-artifact $Executable --format json | Set-Content (Join-Path $OutputDirectory 'artifact.json') -Encoding utf8
if($LASTEXITCODE -ne 0){throw 'REA artifact inspection failed'}
& $NodePath $ReaEntryPath inspect-managed-members $Executable --type NAVFIT98A.PrintModule --format json | Set-Content (Join-Path $OutputDirectory 'print-members.json') -Encoding utf8
if($LASTEXITCODE -ne 0){throw 'REA print-member inspection failed'}
& $NodePath $ReaEntryPath inspect-managed-members $Executable --type NAVFIT98A.IdentifySettings --type NAVFIT98A.CalculationFunctions --type NAVFIT98A.ProceduresModule --type NAVFIT98A.LoadDropDownListsModule --format json | Set-Content (Join-Path $OutputDirectory 'dependencies.json') -Encoding utf8
if($LASTEXITCODE -ne 0){throw 'REA dependency inspection failed'}
if((Get-FileHash -LiteralPath $Executable).Hash -ne $before){throw 'Input binary changed'}
& $NodePath (Join-Path $PSScriptRoot 'summarize-navfit-print-rea.cjs') $OutputDirectory (Join-Path $OutputDirectory 'summary.json')
if($LASTEXITCODE -ne 0){throw 'REA summary failed'}
Write-Output ('PASS: complete REA Evidence retained in '+$OutputDirectory)
