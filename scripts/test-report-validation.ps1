param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource,[string]$InstructionPath,[string]$InstructionTextPath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-rules' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-readiness' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('InstallSource')){if($InstallSource){$caseArgs+='-InstallSource'}}
Invoke-CaseScript 'test-ux-validation' $caseArgs
if($InstructionPath){
$caseArgs=@()
if($PSBoundParameters.ContainsKey('InstructionPath')){$caseArgs+=@('-InstructionPath',$InstructionPath)}
if($PSBoundParameters.ContainsKey('InstructionTextPath')){$caseArgs+=@('-InstructionTextPath',$InstructionTextPath)}
Invoke-CaseScript 'verify-validation-policy' $caseArgs
}else{Write-Output 'SKIP: verify-validation-policy needs -InstructionPath'}
Complete-SuiteRun 'report-validation' $BuildPath $buildHash
