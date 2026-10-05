param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/promotion-status-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'PromotionStatusHarness',(Join-Path $root 'tests/PromotionStatusHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestPromotionStatus()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    $app.CloseCurrentDatabase(); $app.Quit(2); $app=$null
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    if(-not $app.Eval('VerifyPromotionStatusRecovery()')){throw 'New-session promotion status recovery failed'}
    $app.CloseCurrentDatabase()
    @{Status=$result;RecoveryInNewSession=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/promotion-status-results.json')
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
