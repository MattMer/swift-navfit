param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$TaskOrders,[switch]$InstallSource,[switch]$GroupListBenchmark,[string]$BaselinePath,[int]$Groups,[int]$Members,[int]$Trials,[string]$DataSourcePath,[switch]$AdversarialPreviews)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-group-correctness' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('TaskOrders')){if($TaskOrders){$caseArgs+='-TaskOrders'}}
if($PSBoundParameters.ContainsKey('InstallSource')){if($InstallSource){$caseArgs+='-InstallSource'}}
Invoke-CaseScript 'test-group-exit' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('InstallSource')){if($InstallSource){$caseArgs+='-InstallSource'}}
Invoke-CaseScript 'test-live-members' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-native-group-discovery' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('GroupListBenchmark')){if($GroupListBenchmark){$caseArgs+='-GroupListBenchmark'}}
if($PSBoundParameters.ContainsKey('BaselinePath')){$caseArgs+=@('-BaselinePath',$BaselinePath)}
if($PSBoundParameters.ContainsKey('InstallSource')){if($InstallSource){$caseArgs+='-InstallSource'}}
if($PSBoundParameters.ContainsKey('Groups')){$caseArgs+=@('-Groups',$Groups)}
if($PSBoundParameters.ContainsKey('Members')){$caseArgs+=@('-Members',$Members)}
if($PSBoundParameters.ContainsKey('Trials')){$caseArgs+=@('-Trials',$Trials)}
if($PSBoundParameters.ContainsKey('DataSourcePath')){$caseArgs+=@('-DataSourcePath',$DataSourcePath)}
Invoke-CaseScript 'test-ux-group-lifecycle' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-group-reads' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-member-grid' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('AdversarialPreviews')){if($AdversarialPreviews){$caseArgs+='-AdversarialPreviews'}}
Invoke-CaseScript 'test-ux-native-groups' $caseArgs
Complete-SuiteRun 'summary-groups' $BuildPath $buildHash
