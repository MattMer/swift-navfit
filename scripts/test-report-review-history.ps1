param([Parameter(Mandatory=$true)][string]$BuildPath,[ValidateSet('core','extended')][string]$Tier='core',[switch]$SeedDemo,[string]$FixturePath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$caseArgs=@('-BuildPath',$BuildPath)
if($SeedDemo){$caseArgs+='-SeedDemo'}
Invoke-CaseScript 'test-report-tools' $caseArgs
if($Tier -eq 'extended'){
$caseArgs=@('-BuildPath',$BuildPath)
if($FixturePath){$caseArgs+=@('-FixturePath',$FixturePath)}
Invoke-CaseScript 'test-report-notes' $caseArgs
Invoke-CaseScript 'test-revision-compare' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-routing' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-signoff' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-audit-conversations-navigation' @('-BuildPath',$BuildPath)
}
Complete-SuiteRun 'report-review-history' $BuildPath $buildHash
