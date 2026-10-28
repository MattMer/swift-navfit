param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$DisplayFaultOnly)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
Invoke-CaseScript 'test-provenance' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-resource-tools' @('-BuildPath',$BuildPath)
$caseArgs=@('-BuildPath',$BuildPath)
if($DisplayFaultOnly){$caseArgs+='-DisplayFaultOnly'}
Invoke-CaseScript 'test-ux-organization' $caseArgs
Invoke-CaseScript 'test-ux-export-names' @('-BuildPath',$BuildPath)
Complete-SuiteRun 'workspace-resources' $BuildPath $buildHash
