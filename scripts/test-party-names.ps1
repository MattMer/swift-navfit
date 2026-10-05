param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/party-names-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'PartyNamesHarness',(Join-Path $root 'tests/PartyNamesHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestPartyNames()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    [void]$app.Eval('OpenScreen("frmEditor")')
    $app.LoadFromText(5,'FormBoundsHarness',(Join-Path $root 'tests/FormBoundsHarness.bas'))
    [void]$app.Eval('TestButtonOverlap("frmEditor","cmdPartyNames")')
    [void]$app.Eval('OpenPartyNames()')
    [void]$app.Eval('TestFormBounds("frmEditor")')
    [void]$app.Eval('TestFormBounds("frmPartyNames")')
    [void]$app.Eval('CancelPartyNames()')
    $app.DoCmd.Close(2,'frmEditor',2)
    $app.CloseCurrentDatabase(); $app.Quit(2); $app=$null
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    if(-not $app.Eval('VerifyPartyNamesRecovery()')){throw 'New-session structured name recovery failed'}
    $app.CloseCurrentDatabase()
    @{Status=$result;LayoutChecks=3;RecoveryInNewSession=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/party-names-results.json')
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
