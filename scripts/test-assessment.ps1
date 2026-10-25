param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'test-contributor-context.ps1')
$scratch=Join-Path $root ('build/tests/assessment-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    SelectSyntheticContributor $app
    $app.LoadFromText(5,'AssessmentHarness',(Join-Path $root 'tests/AssessmentHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestAssessment()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    $app.CloseCurrentDatabase(); $app.Quit(2); $app=$null
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    SelectSyntheticContributor $app
    if(-not $app.Eval('VerifyAssessmentRecovery()')){throw 'New-session assessment recovery failed'}
    $app.CloseCurrentDatabase()
    @{Status=$result;RecoveryInNewSession=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/assessment-results.json')
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
