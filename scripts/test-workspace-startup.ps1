param([Parameter(Mandatory=$true)][string]$BuildPath,[ValidateSet('core','extended')][string]$Tier='core',[string]$DemoPath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
Invoke-CaseScript 'smoke-test' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-contributor-startup' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-data-link' @('-BuildPath',$BuildPath)
if($Tier -eq 'extended'){Invoke-CaseScript 'test-ux-startup' @('-BuildPath',$BuildPath)}
if($Tier -eq 'extended'){
    if($DemoPath){Invoke-CaseScript 'test-ux-demo' @('-BuildPath',$DemoPath)}
    else{Write-Output 'SKIP: test-ux-demo needs -DemoPath'}
}
Complete-SuiteRun 'workspace-startup' $BuildPath $buildHash
