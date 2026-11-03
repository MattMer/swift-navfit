param([Parameter(Mandatory=$true)][string]$BuildPath,[string]$SourcePath,[switch]$SkipPDF,[string]$SamplePath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-import-commit' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('SourcePath')){$caseArgs+=@('-SourcePath',$SourcePath)}
Invoke-CaseScript 'test-import-review' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-import-rank' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-pdf-import' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-navfit-upgrade' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('SkipPDF')){if($SkipPDF){$caseArgs+='-SkipPDF'}}
Invoke-CaseScript 'test-ux-import-flow' $caseArgs
$caseArgs=@('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-malformed-imports' $caseArgs
#Without -SamplePath the case generates a synthetic NRRM roster.
$caseArgs=@('-BuildPath',$BuildPath)
if($SamplePath){$caseArgs+=@('-SamplePath',$SamplePath)}
Invoke-CaseScript 'test-ruad-import' $caseArgs
Invoke-CaseScript 'test-roster' @('-BuildPath',$BuildPath)
Invoke-CaseScript 'test-reports-due' @('-BuildPath',$BuildPath)
Complete-SuiteRun 'external-import' $BuildPath $buildHash
