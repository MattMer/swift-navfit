param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'test-contributor-context.ps1')
$scratch=Join-Path $root ('build/tests/signoff-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    SelectSyntheticContributor $app
    $app.LoadFromText(5,'SignoffHarness',(Join-Path $root 'tests/SignoffHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestSignoff()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    $app.CloseCurrentDatabase(); $app.Quit(2); $app=$null
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    SelectSyntheticContributor $app
    $app.LoadFromText(5,'SignoffHarness',(Join-Path $root 'tests/SignoffHarness.bas'))
    $recovery=$app.Eval('VerifySignoffRecovery()')
    if(-not $recovery.StartsWith('PASS:')) {throw $recovery}
    $app.CloseCurrentDatabase()
    @{Status=$result;Recovery=$recovery;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/signoff-results.json')
    Write-Output $result
    Write-Output $recovery
} finally {if($null -ne $app){try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2)}}
