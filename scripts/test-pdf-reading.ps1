param([Parameter(Mandatory=$true)][string]$BuildPath,[string]$FixturePath,[string]$LocalFixturePath,[switch]$InstallReader)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
if($LocalFixturePath){
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('LocalFixturePath')){$caseArgs+=@('-FixturePath',$LocalFixturePath)}
if($PSBoundParameters.ContainsKey('InstallReader')){if($InstallReader){$caseArgs+='-InstallReader'}}
Invoke-CaseScript 'test-local-pdf' $caseArgs
}else{Write-Output 'SKIP: test-local-pdf needs -LocalFixturePath'}
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('FixturePath')){$caseArgs+=@('-FixturePath',$FixturePath)}
Invoke-CaseScript 'test-pdf-expanded' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-pdf-counseling-status' $caseArgs
Complete-SuiteRun 'pdf-reading' $BuildPath $buildHash
