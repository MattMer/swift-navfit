param([Parameter(Mandatory=$true)][string]$BuildPath,[string]$BlankPath,[string]$DemoPath,[string]$Candidate,[int]$VerifiedAssertions,[string[]]$VerifiedSuites)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-recursive-defaults' $caseArgs
if($BlankPath -and $DemoPath){
$caseParameters=@{BlankPath=$BlankPath;DemoPath=$DemoPath}
foreach($name in @('Candidate','VerifiedAssertions','VerifiedSuites')){
    if($PSBoundParameters.ContainsKey($name)){$caseParameters[$name]=$PSBoundParameters[$name]}
}
Invoke-CaseScript 'verify-ux-immediate-defaults' -NamedArguments $caseParameters
}else{Write-Output 'SKIP: verify-ux-immediate-defaults needs -BlankPath -DemoPath'}
Complete-SuiteRun 'field-defaults' $BuildPath $buildHash
