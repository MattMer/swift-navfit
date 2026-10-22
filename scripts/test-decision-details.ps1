param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$PersistenceOnly)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'test-contributor-context.ps1')
$scratch=Join-Path $root ('build/tests/decision-details-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    SelectSyntheticContributor $app
    $app.LoadFromText(5,'DecisionDetailsHarness',(Join-Path $root 'tests/DecisionDetailsHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestDecisionDetails()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    if(-not $PersistenceOnly){
    [void]$app.Eval('OpenScreen("frmEditor")')
    $app.LoadFromText(5,'FormBoundsHarness',(Join-Path $root 'tests/FormBoundsHarness.bas'))
    [void]$app.Eval('TestButtonOverlap("frmEditor","cmdDecisionDetails")')
    [void]$app.Eval('OpenDecisionDetails()')
    [void]$app.Eval('TestFormBounds("frmEditor")')
    [void]$app.Eval('TestFormBounds("frmDecisionDetails")')
    [void]$app.Eval('CancelDecisionDetails()')
    $app.DoCmd.Close(2,'frmEditor',2)
    }
    $app.CloseCurrentDatabase(); $app.Quit(2); $app=$null
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    SelectSyntheticContributor $app
    if(-not $app.Eval('VerifyDecisionDetailsRecovery()')){throw 'New-session decision details recovery failed'}
    $app.CloseCurrentDatabase()
    @{Status=$result;LayoutChecks=$(if($PersistenceOnly){0}else{3});LegacyButtonLayoutSkipped=[bool]$PersistenceOnly;RecoveryInNewSession=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/decision-details-results.json')
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
