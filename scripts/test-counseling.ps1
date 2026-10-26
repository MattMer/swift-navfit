param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'test-contributor-context.ps1')
$scratch=Join-Path $root ('build/tests/counseling-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    SelectSyntheticContributor $app
    $app.LoadFromText(5,'CounselingHarness',(Join-Path $root 'tests/CounselingHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestCounseling()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    $app.CloseCurrentDatabase(); $app.Quit(2); $app=$null
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    SelectSyntheticContributor $app
    $app.LoadFromText(5,'CounselingHarness',(Join-Path $root 'tests/CounselingHarness.bas'))
    $recovery=$app.Eval('VerifyCounselingRecovery()')
    if(-not $recovery.StartsWith('PASS:')) {throw $recovery}
    $app.CloseCurrentDatabase()
    @{Status=$result;Recovery=$recovery;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/counseling-results.json')
    Write-Output $result
    Write-Output $recovery
} finally {if($null -ne $app){try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2)}}
