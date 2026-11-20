param([Parameter(Mandatory=$true)][string]$BuildPath)
# Layout audit (tests/LayoutAuditHarness.bas): main windows in realistic states and
# sizes, checked for controls outside the window, overlaps, captions that do not fit,
# windows taller than their space and buttons with the wrong colors. Any finding
# fails, as does missing a defect the harness plants to prove the checks work.
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/layout-audit-'+[guid]::NewGuid()+'.accdb')
$report=[IO.Path]::ChangeExtension($scratch,'.findings.txt')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try{
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'LayoutAuditHarness',(Join-Path $root 'tests/LayoutAuditHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=[string]$app.Eval('TestLayoutAudit("'+$report+'")')
    if(Test-Path -LiteralPath $report){Get-Content -LiteralPath $report|ForEach-Object {Write-Output ('  '+$_)}}
    if(-not $result.StartsWith('PASS:')){throw $result}
    Write-Output $result
}finally{if($null -ne $app){try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
