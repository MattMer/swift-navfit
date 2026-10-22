param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$sourceHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$scratch=Join-Path $root ('build/tests/provenance-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
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
    @{Status=$result;Reopened=$true;BuildSHA256=$sourceHash;TestDatabase=$scratch;Fixture='Self-contained synthetic viewer data; no external import fixture'} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/provenance-results.json')
    Write-Output $result
} finally {
    if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $sourceHash){throw 'Provenance test changed its source artifact'}
}
