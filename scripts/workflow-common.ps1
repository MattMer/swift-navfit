# Shared process boundary: native Access suites run serially in Windows PowerShell.
# A failed child cannot be mistaken for a successful task or receipt.
function InvokeSwiftWorkflowStep([string]$Name,[string[]]$Arguments,[string]$LogFolder){
    $log=Join-Path $LogFolder ('{0:D2}-{1}.log' -f $script:workflowResults.Count,$Name)
    Write-Output ('Running '+$Name+'; log: '+$log)
    $timer=[Diagnostics.Stopwatch]::StartNew()
    try {
        $ErrorActionPreference='Continue'
        & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot ($Name+'.ps1')) @Arguments *> $log
        $stepExit=$LASTEXITCODE
    } finally { $ErrorActionPreference='Stop'; $timer.Stop() }
    if($stepExit -ne 0){
        Get-Content -LiteralPath $log -Tail 35
        throw ($Name+' failed (exit '+$stepExit+'). See '+$log)
    }
    Get-Content -LiteralPath $log | Where-Object {$_ -like 'PASS*'} | Select-Object -Last 3
    $script:workflowResults.Add([pscustomobject]@{Script=$Name;Status='PASS';Seconds=$timer.Elapsed.TotalSeconds;Log=$log})
}

function GetSwiftWorkflowArtifact([string]$Path){
    $item=Get-Item -LiteralPath $Path -ErrorAction Stop
    if($item.PSIsContainer -or $item.Extension -ne '.accdb'){throw 'Expected an existing .accdb file'}
    [pscustomobject]@{Path=$item.FullName;Bytes=$item.Length;SHA256=(Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256).Hash}
}

# Build receipts cover compile/reopen, installed source and form bindings only.
# They never replace a current blank inventory or workflow/migration checks.
function TestSwiftBuildReceipt($Artifact,$Plan){
    try {
        AssertSwiftBuildInputs $Plan
        $receipt=Get-Content -LiteralPath ($Artifact.Path+'.build.json') -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if($receipt.Status -cne 'Built and structurally verified; workflow tests required'){return $false}
        if([IO.Path]::GetFullPath($receipt.Output) -ne $Artifact.Path -or $receipt.SHA256 -ne $Artifact.SHA256){return $false}
        if($receipt.AppVersion -cne $Plan.AppVersion -or $receipt.SchemaVersion -ne $Plan.SchemaVersion){return $false}
        $verification=$receipt.Verification
        foreach($flag in @('Identity','NativeUnloadGuards','RetiredFormsAbsent')){
            if($verification.$flag -isnot [bool] -or -not $verification.$flag){return $false}
        }
        if($verification.SourceModules -ne $Plan.Manifest.sourceModules.Count -or $verification.Forms -le 0){return $false}
        $inputs=@($receipt.Inputs.PSObject.Properties)
        if($inputs.Count -ne $Plan.InputHashes.Count){return $false}
        foreach($key in $Plan.InputHashes.Keys){
            $property=$receipt.Inputs.PSObject.Properties[$key]
            if($null -eq $property -or $property.Value -ne $Plan.InputHashes[$key]){return $false}
        }
        return $true
    } catch {
        # A receipt is an optimization, never permission to bypass verification.
        return $false
    }
}
