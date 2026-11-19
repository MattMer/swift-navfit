param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-recursive-defaults' $caseArgs
Complete-SuiteRun 'field-defaults' $BuildPath $buildHash
