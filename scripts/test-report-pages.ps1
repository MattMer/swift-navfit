param([Parameter(Mandatory=$true)][string]$BuildPath,[ValidateSet('core','extended')][string]$Tier='core',[switch]$EntryLabelsOnly,[switch]$InstallSource,[switch]$EditorReset)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$caseArgs=@('-BuildPath',$BuildPath)
if($EntryLabelsOnly){$caseArgs+='-EntryLabelsOnly'}
if($InstallSource){$caseArgs+='-InstallSource'}
if($EditorReset){$caseArgs+='-EditorReset'}
Invoke-CaseScript 'test-compat-pages' $caseArgs
if($Tier -eq 'extended'){
Invoke-CaseScript 'test-page-switch' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-page-width' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-pending-cells' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-native-dates' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-date-pickers' @('-BuildPath',$BuildPath)
}
Complete-SuiteRun 'report-pages' $BuildPath $buildHash
