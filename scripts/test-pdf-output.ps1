param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource,[switch]$ReusePresentation,[switch]$MetricsOnly,[switch]$ChiefProofOnly,[switch]$ChiefFieldsOnly,[switch]$ChiefCasesOnly,[switch]$TargetComputerOnly,[switch]$TargetComputer,[string]$FixtureDirectory,[string]$ReportPath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
if(-not $TargetComputerOnly){
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('InstallSource')){if($InstallSource){$caseArgs+='-InstallSource'}}
if($PSBoundParameters.ContainsKey('ReusePresentation')){if($ReusePresentation){$caseArgs+='-ReusePresentation'}}
if($PSBoundParameters.ContainsKey('MetricsOnly')){if($MetricsOnly){$caseArgs+='-MetricsOnly'}}
if($PSBoundParameters.ContainsKey('ChiefProofOnly')){if($ChiefProofOnly){$caseArgs+='-ChiefProofOnly'}}
if($PSBoundParameters.ContainsKey('ChiefFieldsOnly')){if($ChiefFieldsOnly){$caseArgs+='-ChiefFieldsOnly'}}
if($PSBoundParameters.ContainsKey('ChiefCasesOnly')){if($ChiefCasesOnly){$caseArgs+='-ChiefCasesOnly'}}
Invoke-CaseScript 'test-pdf-output' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-o3-cells' $caseArgs
}
if($TargetComputer -or $TargetComputerOnly){
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('FixtureDirectory')){$caseArgs+=@('-FixtureDirectory',$FixtureDirectory)}
if($PSBoundParameters.ContainsKey('ReportPath')){$caseArgs+=@('-ReportPath',$ReportPath)}
Invoke-CaseScript 'test-pdf-target-computer' $caseArgs
}else{Write-Output 'SKIP: test-pdf-target-computer needs -TargetComputer (or -TargetComputerOnly)'}
Complete-SuiteRun 'pdf-output' $BuildPath $buildHash
