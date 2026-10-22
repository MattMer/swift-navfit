param(
    [ValidateSet('Preflight','Build','Test','BuildAndTest')][string]$Action='BuildAndTest',
    [string]$BuildPath
)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'build-plan.ps1')
. (Join-Path $PSScriptRoot 'workflow-common.ps1')
$plan=GetSwiftBuildPlan $root (Join-Path $PSScriptRoot 'build-manifest.json')
if($Action -eq 'Test' -and -not $BuildPath){throw '-BuildPath is required for Test'}
if($Action -in @('Build','BuildAndTest')){
    if(-not $BuildPath){$BuildPath=Join-Path $root ('build/dev/Swift_NAVFIT_'+$plan.AppVersion+'_'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'_'+[guid]::NewGuid().ToString('N').Substring(0,8)+'.accdb')}
    $BuildPath=[IO.Path]::GetFullPath($BuildPath)
    if(Test-Path -LiteralPath $BuildPath){throw ('Refusing to overwrite existing database: '+$BuildPath)}
}
$logFolder=Join-Path $root ('build/workflow/'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid())
[void](New-Item -ItemType Directory -Path $logFolder)
$script:workflowResults=[Collections.Generic.List[object]]::new()
InvokeSwiftWorkflowStep 'test-build-plan' @() $logFolder
InvokeSwiftWorkflowStep 'test-vba-source-normalization' @() $logFolder
if($Action -in @('Build','BuildAndTest')){InvokeSwiftWorkflowStep 'build' @('-OutputPath',$BuildPath) $logFolder}
if($Action -in @('Test','BuildAndTest')){
    $before=GetSwiftWorkflowArtifact $BuildPath
    InvokeSwiftWorkflowStep 'verify-build-artifact' @('-BuildPath',$before.Path,'-RequireBlank') $logFolder
    InvokeSwiftWorkflowStep 'test-report-tools' @('-BuildPath',$before.Path,'-SeedDemo') $logFolder
    foreach($suite in @('test-contributor-startup','smoke-test','test-compat-pages','test-ux-workflows','test-directory-merge')){
        InvokeSwiftWorkflowStep $suite @('-BuildPath',$before.Path) $logFolder
    }
    $after=GetSwiftWorkflowArtifact $before.Path
    if($after.SHA256 -ne $before.SHA256){throw 'Development tests changed the supplied artifact'}
}
AssertSwiftBuildInputs $plan
$artifact=if($Action -ne 'Preflight'){GetSwiftWorkflowArtifact $BuildPath}else{$null}
$receipt=Join-Path $logFolder 'verification.json'
[pscustomobject]@{Status='PASS';Action=$Action;AppVersion=$plan.AppVersion;SchemaVersion=$plan.SchemaVersion;Artifact=$artifact;Inputs=$plan.InputHashes;Steps=@($workflowResults.ToArray());Scope='Development checks only; release and physical acceptance remain separate'} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $receipt
Write-Output ('PASS: '+$Action+'; receipt: '+$receipt)
if($artifact){Write-Output ('Artifact: '+$artifact.Path)}
