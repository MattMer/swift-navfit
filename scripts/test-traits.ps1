param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'test-contributor-context.ps1')
$sourceHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$scratch=Join-Path $root ('build/tests/traits-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'TraitsHarness',(Join-Path $root 'tests/TraitsHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestTraits()')
    if($result -notlike 'PASS:*'){throw $result}
    $app.CloseCurrentDatabase();$app.Quit(2)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    SelectSyntheticContributor $app
    if(-not $app.Eval('VerifyTraitRecovery()')){throw 'Trait recovery in new session failed'}
    $app.CloseCurrentDatabase()
    @{Status=$result;RecoveryInNewSession=$true;BuildSHA256=$sourceHash;TestDatabase=$scratch;Scope='Trait persistence, rollback, legacy restore and recovery; catalogue/UI checked separately'} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/traits-results.json')
    Write-Output $result
    Write-Output 'PASS: invalid pending trait recovery in a new Access session'
} finally {
    if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $sourceHash){throw 'Trait test changed its source artifact'}
}
