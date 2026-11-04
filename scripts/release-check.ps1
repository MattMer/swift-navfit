param(
    [Parameter(Mandatory=$true)][string]$BlankPath,
    [Parameter(Mandatory=$true)][string]$DemoPath,
    [switch]$CurrentFixturesOnly
)
# Verification only. Publication remains subject to the CONTRIBUTING.md delivery gates.
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'build-plan.ps1')
. (Join-Path $PSScriptRoot 'workflow-common.ps1')
. (Join-Path $PSScriptRoot 'build-retention.ps1')
$plan=GetSwiftBuildPlan $root (Join-Path $PSScriptRoot 'build-manifest.json')
$blank=GetSwiftWorkflowArtifact $BlankPath
$demo=GetSwiftWorkflowArtifact $DemoPath
if($blank.Path -eq $demo.Path){throw 'Blank and Demo must be distinct files'}
[void](InvokeSwiftScratchCleanup -Root $root -Apply -PreservePath @($blank.Path,$demo.Path))
$scratchRun=StartSwiftScratchRun $root
$scratchStatus='FAIL'
try {
$logFolder=Join-Path $root ('build/workflow/release-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid())
[void](New-Item -ItemType Directory -Path $logFolder)
$script:workflowResults=[Collections.Generic.List[object]]::new()
InvokeSwiftWorkflowStep 'test-source-contracts' @() $logFolder
InvokeSwiftWorkflowStep 'verify-build-artifact' @('-BuildPath',$blank.Path,'-RequireBlank') $logFolder
InvokeSwiftWorkflowStep 'verify-build-artifact' @('-BuildPath',$demo.Path) $logFolder
InvokeSwiftWorkflowStep 'test-workspace-startup' @('-BuildPath',$blank.Path,'-Tier','extended','-DemoPath',$demo.Path) $logFolder
InvokeSwiftWorkflowStep 'test-report-navigation' @('-BuildPath',$blank.Path) $logFolder
#Navigation cases expect a blank workspace; the Demo gets its own startup and launch checks.
InvokeSwiftWorkflowStep 'cases/test-ux-startup' @('-BuildPath',$demo.Path) $logFolder
InvokeSwiftWorkflowStep 'cases/test-preview-launch' @('-BuildPath',$demo.Path) $logFolder
$directoryArguments=@('-Tier','release','-BlankPath',$blank.Path,'-DemoPath',$demo.Path)
if($CurrentFixturesOnly){$directoryArguments+='-CurrentFixturesOnly'}
InvokeSwiftWorkflowStep 'test-directory-management' $directoryArguments $logFolder
$migration=Get-Content (Join-Path $root 'build/r115-data-import-verification.json') -Raw | ConvertFrom-Json
$expectedSchemas=if($CurrentFixturesOnly){@($plan.SchemaVersion)}else{@(1..$plan.SchemaVersion)}
if($migration.Status -ne 'PASS' -or (Compare-Object $expectedSchemas @($migration.Fixtures|ForEach-Object {[int]$_.Schema}|Sort-Object -Unique))){throw 'Requested schema coverage is incomplete'}
foreach($guard in @('Rollback','ChangedSourceRejected','PopulatedTargetRejected','UnknownTableRejected','FutureSchemaRejected','UIEntryPointTested')){
    if($migration.$guard -ne $true){throw ('Missing release migration guard: '+$guard)}
}
foreach($suite in @('test-workspace-startup','test-report-pages','test-work-review','test-summary-groups')){
    InvokeSwiftWorkflowStep $suite @('-BuildPath',$blank.Path) $logFolder
}
InvokeSwiftWorkflowStep 'test-import-engine' @('-BuildPath',$blank.Path,'-Tier','release','-SourcePath',$demo.Path) $logFolder
InvokeSwiftWorkflowStep 'develop' @('-Action','ExtendedTest','-BuildPath',$blank.Path) $logFolder
InvokeSwiftWorkflowStep 'test-workspace-migration' @('-BuildPath',$blank.Path,'-Cases','test-report-message-import','-DemoPath',$demo.Path) $logFolder
foreach($artifact in @($blank,$demo)){
    if((GetSwiftWorkflowArtifact $artifact.Path).SHA256 -ne $artifact.SHA256){throw ('Release checks changed '+$artifact.Path)}
}
AssertSwiftBuildInputs $plan
$receipt=Join-Path $logFolder 'verification.json'
[pscustomobject]@{Status='PASS';AppVersion=$plan.AppVersion;SchemaVersion=$plan.SchemaVersion;Blank=$blank;Demo=$demo;Inputs=$plan.InputHashes;HistoricalFixturesRetested=(-not $CurrentFixturesOnly);HistoricalRetestWaiver=$(if($CurrentFixturesOnly){'User explicitly waived old-output fixture testing for this delivery on 2026-10-21'}else{$null});Migration=$migration;Steps=@($workflowResults.ToArray());Scope='Native release regression for requested fixtures, populated transfer and rollback guards; excludes physical acceptance and publication'} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $receipt
Write-Output ('PASS: release verification; receipt: '+$receipt)
$scratchStatus='PASS'
} finally {
    try{CompleteSwiftScratchRun $scratchRun $scratchStatus}catch{Write-Warning ('Scratch retention record failed: '+$_.Exception.Message)}
}
