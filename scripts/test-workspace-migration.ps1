param([Parameter(Mandatory=$true)][string]$BuildPath,[ValidateSet('test-swift-data-import','test-workspace-upgrade','test-report-message-import','test-directory-merge-transfer')][string[]]$Cases=@('test-swift-data-import','test-workspace-upgrade','test-report-message-import','test-directory-merge-transfer'),[string]$Schemas,[string]$CurrentSchemaPath,[switch]$CurrentFixturesOnly,[switch]$SkipKeptReleases,[string]$FixturePath,[string]$Schema20Path,[Alias('Schema21Path')][string]$SourceWorkspacePath,[string]$DemoPath,[string]$SourcePath)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
if($Cases -contains 'test-swift-data-import'){
$caseArgs=@('-BuildPath',$BuildPath)
if($PSBoundParameters.ContainsKey('Schemas')){$caseArgs+=@('-Schemas',$Schemas)}
if($PSBoundParameters.ContainsKey('CurrentSchemaPath')){$caseArgs+=@('-CurrentSchemaPath',$CurrentSchemaPath)}
if($PSBoundParameters.ContainsKey('FixturePath')){$caseArgs+=@('-FixturePath',$FixturePath)}
if($PSBoundParameters.ContainsKey('CurrentFixturesOnly')){if($CurrentFixturesOnly){$caseArgs+='-CurrentFixturesOnly'}}
if($PSBoundParameters.ContainsKey('SkipKeptReleases')){if($SkipKeptReleases){$caseArgs+='-SkipKeptReleases'}}
Invoke-CaseScript 'test-swift-data-import' $caseArgs
}
if($Cases -contains 'test-workspace-upgrade'){
$caseArgs=@('-TemplatePath',$BuildPath)
if($PSBoundParameters.ContainsKey('Schema20Path')){$caseArgs+=@('-Schema20Path',$Schema20Path)}
if($PSBoundParameters.ContainsKey('SourceWorkspacePath')){$caseArgs+=@('-SourceWorkspacePath',$SourceWorkspacePath)}
Invoke-CaseScript 'test-workspace-upgrade' $caseArgs
}
if($Cases -contains 'test-report-message-import'){
if($DemoPath){
$caseArgs=@('-BlankPath',$BuildPath,'-DemoPath',$DemoPath)
Invoke-CaseScript 'test-report-message-import' $caseArgs
}else{Write-Output 'SKIP: test-report-message-import needs -DemoPath'}
}
if($Cases -contains 'test-directory-merge-transfer'){
if($SourcePath){
$caseArgs=@('-TemplatePath',$BuildPath,'-SourcePath',$SourcePath)
Invoke-CaseScript 'test-directory-merge-transfer' $caseArgs
}else{Write-Output 'SKIP: test-directory-merge-transfer needs -SourcePath'}
}
Complete-SuiteRun 'workspace-migration' $BuildPath $buildHash
