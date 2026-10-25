param(
    [ValidateSet('Preflight','Build','Test','BuildAndTest','Compile','ExtendedTest')][string]$Action='BuildAndTest',
    [string]$BuildPath,
    [string]$CompiledPath
)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'build-plan.ps1')
. (Join-Path $PSScriptRoot 'workflow-common.ps1')
$plan=GetSwiftBuildPlan $root (Join-Path $PSScriptRoot 'build-manifest.json')
if($Action -in @('Test','Compile','ExtendedTest') -and -not $BuildPath){throw '-BuildPath is required for Test, Compile or ExtendedTest'}
if($CompiledPath -and $Action -ne 'Compile'){throw '-CompiledPath is only valid for Compile'}
if($CompiledPath){
    $CompiledPath=[IO.Path]::GetFullPath($CompiledPath)
    if([IO.Path]::GetExtension($CompiledPath) -ne '.accde'){throw 'Compiled output must use .accde'}
    if(Test-Path -LiteralPath $CompiledPath){throw ('Refusing to overwrite existing output: '+$CompiledPath)}
}
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
InvokeSwiftWorkflowStep 'test-field-definitions' @() $logFolder
$compiledArtifact=$null;$compiler=$null
if($Action -eq 'Compile'){
    $before=GetSwiftWorkflowArtifact $BuildPath
    InvokeSwiftWorkflowStep 'verify-build-artifact' @('-BuildPath',$before.Path,'-RequireBlank') $logFolder
    $copy=Join-Path $logFolder 'compile-source.accdb'
    $candidate=Join-Path $logFolder 'candidate.accde'
    Copy-Item -LiteralPath $before.Path -Destination $copy
    $app=$null;$compilerCleanupFailed=$false
    try {
        $app=New-Object -ComObject Access.Application
        $app.AutomationSecurity=1
        $executable=Join-Path ([string]$app.SysCmd(9)) 'MSACCESS.EXE'
        $reader=[IO.BinaryReader]::new([IO.File]::OpenRead($executable))
        try {
            $reader.BaseStream.Position=0x3c;$pe=$reader.ReadInt32()
            $reader.BaseStream.Position=$pe
            if($reader.ReadUInt32() -ne 0x4550){throw 'Invalid Access PE signature'}
            $machine=$reader.ReadUInt16()
            $architecture=switch($machine){0x8664{'x64'} 0x14c{'x86'} default{throw ('Untested Access architecture: '+$machine)}}
        } finally {$reader.Dispose()}
        $compiler=[pscustomobject]@{Version=[string]$app.Version;FileVersion=(Get-Item -LiteralPath $executable).VersionInfo.FileVersion;Architecture=$architecture;Executable=$executable;Method='Access SysCmd 603 (undocumented; exact-host verification required)'}
        Write-Output ('Controller: development ACCDE compilation; disposable source: '+$copy+'; Access '+$compiler.FileVersion+' '+$architecture)
        # No source import or parallel assembly: convert the verified source-only build.
        # 603 is not a documented SysCmd action. Keep it isolated, fail closed and
        # verify the resulting compiled binary through real workflows on this host.
        # Windows PowerShell's early-bound interop rejects raw integers outside
        # the published enum; preserve the native action as an enum instance.
        $compileAction=[Enum]::ToObject([Microsoft.Office.Interop.Access.AcSysCmdAction],603)
        [void]$app.SysCmd($compileAction,$copy,$candidate)
        if(-not (Test-Path -LiteralPath $candidate -PathType Leaf)){throw 'Access did not create an ACCDE candidate'}
    } catch { $primary=$_;throw $primary }
    finally {
        if($app){
            try{$app.Quit(2)}catch{$compilerCleanupFailed=$true;Write-Warning ('Owned compiler cleanup: '+$_.Exception.Message)}
            try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{Write-Warning ('Owned compiler release: '+$_.Exception.Message)}
        }
    }
    if($compilerCleanupFailed){throw 'Compiler session did not close normally; no test-controller handoff or accepted output'}
    $checkerHashes=@{Develop=(Get-FileHash -LiteralPath $PSCommandPath).Hash;CompiledTests=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'test-compiled-distribution.ps1')).Hash}
    InvokeSwiftWorkflowStep 'test-compiled-distribution' @('-BuildPath',$before.Path,'-CompiledPath',$candidate) $logFolder
    if((Get-FileHash -LiteralPath $PSCommandPath).Hash -ne $checkerHashes.Develop -or (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'test-compiled-distribution.ps1')).Hash -ne $checkerHashes.CompiledTests){throw 'Compiled verification code changed during the run'}
    AssertSwiftBuildInputs $plan
    if((GetSwiftWorkflowArtifact $before.Path).SHA256 -ne $before.SHA256){throw 'ACCDE compilation changed the source artifact'}
    if(-not $CompiledPath){$CompiledPath=Join-Path $root ('build/dev/Swift_NAVFIT_'+$plan.AppVersion+'_'+$architecture+'_'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'_'+[guid]::NewGuid().ToString('N').Substring(0,8)+'.accde')}
    [void](New-Item -ItemType Directory -Force -Path (Split-Path $CompiledPath -Parent))
    $candidateHash=(Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash
    [IO.File]::Copy($candidate,$CompiledPath,$false)
    if((Get-FileHash -LiteralPath $CompiledPath -Algorithm SHA256).Hash -ne $candidateHash){throw 'Accepted output differs from tested ACCDE candidate'}
    $compiledArtifact=[pscustomobject]@{Path=$CompiledPath;Bytes=(Get-Item -LiteralPath $CompiledPath).Length;SHA256=$candidateHash;Signature='Unsigned; approved certificate required';NormalSecurityAccepted=$false;Compiler=$compiler;Source=$before;Checkers=$checkerHashes}
}
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
if($Action -eq 'ExtendedTest'){
    #The extended tier in tests/runners.json, run weekly and before every release.
    #Every runner runs even after a failure, so one run reports the whole tier.
    $before=GetSwiftWorkflowArtifact $BuildPath
    InvokeSwiftWorkflowStep 'verify-build-artifact' @('-BuildPath',$before.Path,'-RequireBlank') $logFolder
    $extended=@((Get-Content (Join-Path $root 'tests/runners.json') -Raw|ConvertFrom-Json).runners|Where-Object {$_.tier -eq 'extended'})
    $failed=@()
    foreach($runner in $extended){
        try{InvokeSwiftWorkflowStep ([string]$runner.name) (@('-BuildPath',$before.Path)+@($runner.extraArgs|Where-Object {$_})) $logFolder}catch{$failed+=[string]$runner.name;Write-Output $_.Exception.Message}
    }
    if((GetSwiftWorkflowArtifact $before.Path).SHA256 -ne $before.SHA256){throw 'Extended tests changed the supplied artifact'}
    if($failed.Count){throw ('Extended tier: '+$failed.Count+' of '+$extended.Count+' runners failed: '+($failed -join ', ')+'. Logs: '+$logFolder)}
}
AssertSwiftBuildInputs $plan
$artifact=if($Action -ne 'Preflight'){GetSwiftWorkflowArtifact $BuildPath}else{$null}
$receipt=Join-Path $logFolder 'verification.json'
[pscustomobject]@{Status='PASS';Action=$Action;AppVersion=$plan.AppVersion;SchemaVersion=$plan.SchemaVersion;Artifact=$artifact;CompiledArtifact=$compiledArtifact;Inputs=$plan.InputHashes;Steps=@($workflowResults.ToArray());Scope='Development checks only; release and physical acceptance remain separate'} | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $receipt
Write-Output ('PASS: '+$Action+'; receipt: '+$receipt)
if($artifact){Write-Output ('Artifact: '+$artifact.Path)}
if($compiledArtifact){Write-Output ('Compiled artifact: '+$compiledArtifact.Path)}
