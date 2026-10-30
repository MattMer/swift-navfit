param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource,[switch]$VerifyFieldCoverage,[string]$OutputPath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('InstallSource')){if($InstallSource){$caseArgs+='-InstallSource'}}
Invoke-CaseScript 'test-navfit-adapter' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('InstallSource')){if($InstallSource){$caseArgs+='-InstallSource'}}
Invoke-CaseScript 'test-navfit-group-export' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('InstallSource')){if($InstallSource){$caseArgs+='-InstallSource'}}
Invoke-CaseScript 'test-ux-navfit' $caseArgs
if($VerifyFieldCoverage){
$caseArgs=@()
if($PSBoundParameters.ContainsKey('OutputPath')){$caseArgs+=@('-OutputPath',$OutputPath)}
Invoke-CaseScript 'verify-navfit-field-coverage' $caseArgs
}else{Write-Output 'SKIP: verify-navfit-field-coverage needs -VerifyFieldCoverage'}
Complete-SuiteRun 'navfit-exchange' $BuildPath $buildHash
