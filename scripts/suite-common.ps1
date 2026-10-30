# Shared case runner for feature test suites (dot-source it).
# Every case runs, even after an earlier one fails, on its own disposable copy, so one
# suite run reports the whole feature. Complete-SuiteRun throws if any case failed.
$script:suiteFailures=@()
$script:suiteCases=0

# Run one case body; record a failure and keep going.
function Invoke-SuiteCase([string]$Name,[scriptblock]$Body){
    $script:suiteCases++
    try{& $Body}catch{
        $script:suiteFailures+=($Name+': '+$_.Exception.Message)
        Write-Output ('FAIL: '+$Name+': '+$_.Exception.Message)
    }
}

# The common harness case: copy the build, load one tests/*.bas module, compile,
# evaluate its entry point and require a PASS result. $Before runs with the open
# application before the module loads; $After runs before it quits (best-effort UI
# cleanup). $SourceModules reloads named src/ modules into the copy first.
function Invoke-HarnessCase([string]$Name,[string]$BuildPath,[string]$Module,[string]$Entry,[bool]$Exclusive=$true,[scriptblock]$Before=$null,[scriptblock]$After=$null,[string[]]$SourceModules=@()){
    Invoke-SuiteCase $Name {
        $root=Split-Path $PSScriptRoot -Parent
        $scratch=Join-Path $root ('build/tests/'+$Name+'-'+[guid]::NewGuid()+'.accdb')
        Copy-Item -LiteralPath $BuildPath -Destination $scratch
        $app=New-Object -ComObject Access.Application
        try{
            $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$Exclusive)
            if($Before){& $Before $app}
            foreach($source in $SourceModules){$app.LoadFromText(5,$source,(Join-Path $root ('src/'+$source+'.bas')))}
            $app.LoadFromText(5,$Module,(Join-Path $root ('tests/'+$Module+'.bas')))
            $app.DoCmd.RunCommand(126)
            $result=[string]$app.Eval($Entry);Write-Output $result
            if($result -notlike 'PASS:*'){throw $result}
        }finally{
            if($After){try{& $After $app}catch{}}
            $app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
        }
    }
}

# Run one case script from scripts/cases in its own Windows PowerShell process.
function Invoke-CaseScript([string]$Name,[string[]]$Arguments=@(),[hashtable]$NamedArguments=@{}){
    Invoke-SuiteCase $Name {
        $casePath=Join-Path $PSScriptRoot ('cases/'+$Name+'.ps1')
        if($NamedArguments.Count){
            # Windows PowerShell -File cannot pass array-valued parameters. Keep
            # their types through built-in CLIXML and splat them in the child.
            if($Arguments.Count){throw 'Use either Arguments or NamedArguments'}
            $payload=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes([Management.Automation.PSSerializer]::Serialize($NamedArguments)))
            $command='$caseParameters=[Management.Automation.PSSerializer]::Deserialize([Text.Encoding]::Unicode.GetString([Convert]::FromBase64String("'+$payload+'"))); & '+"'"+$casePath.Replace("'","''")+"'"+' @caseParameters'
            $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
            & powershell.exe -NoProfile -EncodedCommand $encoded
        }else{
            & powershell.exe -NoProfile -File $casePath @Arguments
        }
        if($LASTEXITCODE -ne 0){throw ($Name+' exited '+$LASTEXITCODE)}
    }
}

# Fail the suite if any case failed, after confirming the build was not changed.
function Complete-SuiteRun([string]$Suite,[string]$BuildPath,[string]$BuildHash){
    if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $BuildHash){$script:suiteFailures+='suite changed its source artifact'}
    if($script:suiteFailures.Count){throw ($Suite+': '+$script:suiteFailures.Count+' of '+$script:suiteCases+' cases failed: '+($script:suiteFailures -join ' | '))}
    Write-Output ('PASS: '+$Suite+' suite, '+$script:suiteCases+' cases')
}
