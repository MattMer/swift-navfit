param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$PersistenceOnly)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-member-details' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('PersistenceOnly')){if($PersistenceOnly){$caseArgs+='-PersistenceOnly'}}
Invoke-CaseScript 'test-party-names' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('PersistenceOnly')){if($PersistenceOnly){$caseArgs+='-PersistenceOnly'}}
Invoke-CaseScript 'test-report-people' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-promotion-status' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-member-psg' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-counseling' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-prior-grid' $caseArgs
Complete-SuiteRun 'member-profiles' $BuildPath $buildHash
