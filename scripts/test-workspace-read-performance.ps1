param([Parameter(Mandatory=$true)][string]$BuildPath,[string]$BaselinePath,[switch]$InstallSource,[int]$Size=500,[int]$Trials=5,[string]$DataSourcePath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
if(-not $BaselinePath){Write-Output 'SKIP: test-batched-reads needs -BaselinePath'}
else{
    $caseArgs=@('-BuildPath',$BuildPath,'-BaselinePath',$BaselinePath,'-Size',[string]$Size,'-Trials',[string]$Trials)
    if($InstallSource){$caseArgs+='-InstallSource'}
    if($DataSourcePath){$caseArgs+=@('-DataSourcePath',$DataSourcePath)}
    Invoke-CaseScript 'test-batched-reads' $caseArgs
}
Invoke-CaseScript 'test-ux-read-rows' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-grid-reads' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-performance' @('-BuildPath',$BuildPath)
Complete-SuiteRun 'workspace-read-performance' $BuildPath $buildHash
