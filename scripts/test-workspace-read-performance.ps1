param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
Invoke-CaseScript 'test-ux-read-rows' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-grid-reads' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-performance' @('-BuildPath',$BuildPath)
Complete-SuiteRun 'workspace-read-performance' $BuildPath $buildHash
