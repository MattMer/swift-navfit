param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$PersistenceOnly)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-administration' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-assessment' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('PersistenceOnly')){if($PersistenceOnly){$caseArgs+='-PersistenceOnly'}}
Invoke-CaseScript 'test-decision-details' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-service-details' $caseArgs
Complete-SuiteRun 'report-field-persistence' $BuildPath $buildHash
