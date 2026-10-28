param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
Invoke-CaseScript 'test-navigation-context' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-review-return' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-shell' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-surface' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-visual' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-context-help' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-preview-launch' @('-BuildPath',$BuildPath)
Complete-SuiteRun 'report-navigation' $BuildPath $buildHash
