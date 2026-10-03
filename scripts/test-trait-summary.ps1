param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/trait-summary-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'TraitSummaryHarness',(Join-Path $root 'tests/TraitSummaryHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestTraitSummary()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    $app.CloseCurrentDatabase()
    @{Status=$result;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/trait-summary-results.json')
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
