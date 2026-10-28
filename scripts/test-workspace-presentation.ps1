param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$UseWorkingSource)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
Invoke-CaseScript 'test-accessibility' @('-BuildPath',$BuildPath)
$caseArgs=@('-BuildPath',$BuildPath)
if($UseWorkingSource){$caseArgs+='-UseWorkingSource'}
Invoke-CaseScript 'test-table-sort' $caseArgs
Invoke-CaseScript 'verify-table-sort-bindings' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-admin-feedback' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-feedback' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-help' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-responsive-identity' @('-BuildPath',$BuildPath)
Complete-SuiteRun 'workspace-presentation' $BuildPath $buildHash
