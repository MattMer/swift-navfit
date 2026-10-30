param([Parameter(Mandatory=$true)][string]$BuildPath,[ValidateSet('core','extended')][string]$Tier='core',[switch]$ChiefComments,[switch]$InstallSource,[string]$LegacyPath,[string]$OlderPath,[string]$PDFPath,[switch]$CrossWorkflowOnly,[switch]$ApplyFaultsOnly,[switch]$SkipUnhandledFailure)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('ChiefComments')){if($ChiefComments){$caseArgs+='-ChiefComments'}}
Invoke-CaseScript 'test-ux-workflows' $caseArgs
if($Tier -eq 'extended'){
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('InstallSource')){if($InstallSource){$caseArgs+='-InstallSource'}}
if($PSBoundParameters.ContainsKey('LegacyPath')){$caseArgs+=@('-LegacyPath',$LegacyPath)}
if($PSBoundParameters.ContainsKey('OlderPath')){$caseArgs+=@('-OlderPath',$OlderPath)}
if($PSBoundParameters.ContainsKey('ChiefComments')){if($ChiefComments){$caseArgs+='-ChiefComments'}}
Invoke-CaseScript 'test-major' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-major-ui' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('InstallSource')){if($InstallSource){$caseArgs+='-InstallSource'}}
Invoke-CaseScript 'test-ux-my-work' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-work-inclusion' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-ux-work-refresh' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('InstallSource')){if($InstallSource){$caseArgs+='-InstallSource'}}
if($PSBoundParameters.ContainsKey('LegacyPath')){$caseArgs+=@('-LegacyPath',$LegacyPath)}
if($PSBoundParameters.ContainsKey('PDFPath')){$caseArgs+=@('-PDFPath',$PDFPath)}
if($PSBoundParameters.ContainsKey('CrossWorkflowOnly')){if($CrossWorkflowOnly){$caseArgs+='-CrossWorkflowOnly'}}
if($PSBoundParameters.ContainsKey('ApplyFaultsOnly')){if($ApplyFaultsOnly){$caseArgs+='-ApplyFaultsOnly'}}
if($PSBoundParameters.ContainsKey('SkipUnhandledFailure')){if($SkipUnhandledFailure){$caseArgs+='-SkipUnhandledFailure'}}
Invoke-CaseScript 'test-wizard-state' $caseArgs
}
Complete-SuiteRun 'work-review' $BuildPath $buildHash
