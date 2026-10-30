param([string]$WorkspacePath,[string]$ReturnedPath,[int]$ReportID,[string]$Marker,[switch]$AlreadyApplied,[string]$BeforePath,[string]$AfterPath,[string]$EvidencePath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
# The native-return case intentionally applies changes to the supplied disposable workspace.
# Hash the read-only returned file (or comparison source), never the workspace it edits.
$hashPath=if($ReturnedPath){$ReturnedPath}elseif($BeforePath){$BeforePath}else{$null}
if($hashPath){$buildHash=(Get-FileHash -LiteralPath $hashPath).Hash}
if($WorkspacePath -and $ReturnedPath -and $PSBoundParameters.ContainsKey('ReportID')){
$caseArgs=@()
if($PSBoundParameters.ContainsKey('WorkspacePath')){$caseArgs+=@('-WorkspacePath',$WorkspacePath)}
if($PSBoundParameters.ContainsKey('ReturnedPath')){$caseArgs+=@('-ReturnedPath',$ReturnedPath)}
if($PSBoundParameters.ContainsKey('ReportID')){$caseArgs+=@('-ReportID',$ReportID)}
if($PSBoundParameters.ContainsKey('Marker')){$caseArgs+=@('-Marker',$Marker)}
if($PSBoundParameters.ContainsKey('AlreadyApplied')){if($AlreadyApplied){$caseArgs+='-AlreadyApplied'}}
Invoke-CaseScript 'test-navfit-native-return' $caseArgs
}else{Write-Output 'SKIP: test-navfit-native-return needs -WorkspacePath -ReturnedPath -ReportID'}
if($BeforePath -and $AfterPath -and $EvidencePath){
$caseArgs=@()
if($PSBoundParameters.ContainsKey('BeforePath')){$caseArgs+=@('-BeforePath',$BeforePath)}
if($PSBoundParameters.ContainsKey('AfterPath')){$caseArgs+=@('-AfterPath',$AfterPath)}
if($PSBoundParameters.ContainsKey('EvidencePath')){$caseArgs+=@('-EvidencePath',$EvidencePath)}
Invoke-CaseScript 'verify-navfit-native-roundtrip' $caseArgs
}else{Write-Output 'SKIP: verify-navfit-native-roundtrip needs -BeforePath -AfterPath -EvidencePath'}
if($hashPath){Complete-SuiteRun 'navfit-native-return' $hashPath $buildHash}else{Write-Output 'PASS: navfit-native-return suite, 0 cases (manual inputs absent)'}
