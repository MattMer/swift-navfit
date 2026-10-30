param([Parameter(Mandatory=$true)][string]$BuildPath,[ValidateSet('extended','release')][string]$Tier='extended',[string]$BaselinePath,[string]$DemoPath,[string]$SourcePath,[switch]$SkipBaseline,[switch]$InstallSource,[switch]$CommitFaults)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
if($Tier -eq 'extended'){
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('BaselinePath')){$caseArgs+=@('-BaselinePath',$BaselinePath)}
if($PSBoundParameters.ContainsKey('DemoPath')){$caseArgs+=@('-DemoPath',$DemoPath)}
Invoke-CaseScript 'test-bulk-workspace-import' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('BaselinePath')){$caseArgs+=@('-BaselinePath',$BaselinePath)}
if($PSBoundParameters.ContainsKey('DemoPath')){$caseArgs+=@('-DemoPath',$DemoPath)}
if($PSBoundParameters.ContainsKey('SourcePath')){$caseArgs+=@('-SourcePath',$SourcePath)}
if($PSBoundParameters.ContainsKey('SkipBaseline')){if($SkipBaseline){$caseArgs+='-SkipBaseline'}}
Invoke-CaseScript 'test-engine-import' $caseArgs
}else{
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('SourcePath')){$caseArgs+=@('-SourcePath',$SourcePath)}
if($PSBoundParameters.ContainsKey('InstallSource')){if($InstallSource){$caseArgs+='-InstallSource'}}
if($PSBoundParameters.ContainsKey('CommitFaults')){if($CommitFaults){$caseArgs+='-CommitFaults'}}
Invoke-CaseScript 'test-engine-verification' $caseArgs
}
Complete-SuiteRun 'import-engine' $BuildPath $buildHash
