param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build\tests\recovery-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'RecoveryHarness',(Join-Path $root 'tests\RecoveryHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $prepared=$app.Eval('PrepareRecovery()');if(-not $prepared.StartsWith('PASS:')){throw $prepared}
    $app.CloseCurrentDatabase();$app.Quit(2);$app=$null
    $preparedCopy=$scratch.Replace('.accdb','-prepared.accdb')
    Copy-Item -LiteralPath $scratch -Destination $preparedCopy
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $verified=$app.Eval('VerifyRecovery()');if(-not $verified.StartsWith('PASS:')){throw $verified}
    [void]$app.Eval('PrepareTimerCheck()')
    Start-Sleep -Milliseconds 700
    if(-not $app.Eval('FinishTimerCheck()')){throw 'Timer check failed.'}
    $app.CloseCurrentDatabase()
    @{Preparation=$prepared;Recovery=$verified;AutomaticTimerAndActiveTextVerified=$true;ReopenedInNewAccessSession=$true;PreparedCheckpointDatabase=$preparedCopy;TestDatabase=$scratch;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash}|ConvertTo-Json|Set-Content (Join-Path $root 'build\tests\recovery-results.json')
    Write-Output $prepared;Write-Output $verified
}finally{if($null -ne $app){$app.Quit(2)}}
