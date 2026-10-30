param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'test-contributor-context.ps1')
$scratch=Join-Path $root ('build/tests/administration-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    SelectSyntheticContributor $app
    $app.LoadFromText(5,'AdministrationHarness',(Join-Path $root 'tests/AdministrationHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestAdministration()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    $app.CloseCurrentDatabase(); $app.Quit(2); $app=$null
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    SelectSyntheticContributor $app
    if(-not $app.Eval('VerifyAdministrativeRecovery()')) {throw 'Administrative recovery in new session failed'}
    $app.CloseCurrentDatabase()
    @{Status=$result;RecoveryInNewSession=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/administration-results.json')
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
