param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
& (Join-Path $PSScriptRoot 'test-tools.ps1') -BuildPath $BuildPath
$fixture=Get-Content (Join-Path $root 'build/tests/tools-results.json') -Raw | ConvertFrom-Json
$scratch=$fixture.TestDatabase
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'ProvenanceHarness',(Join-Path $root 'tests/ProvenanceHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestProvenance()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    $app.CloseCurrentDatabase(); $app.OpenCurrentDatabase($scratch,$true)
    if(-not $app.Eval('OpenImportProvenance(DMin("ImportReceiptID","ImportReceipt"))')) {throw 'Viewer reopen failed'}
    $app.CloseCurrentDatabase()
    @{Status=$result;Reopened=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/provenance-results.json')
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
