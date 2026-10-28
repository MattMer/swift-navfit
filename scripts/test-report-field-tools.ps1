param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
Invoke-CaseScript 'test-report-context' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-block-tools' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-header-review' @('-BuildPath',$BuildPath)
Complete-SuiteRun 'report-field-tools' $BuildPath $buildHash
