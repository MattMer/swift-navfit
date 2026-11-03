param([Parameter(Mandatory=$true)][string]$BuildPath)
#F03 reports due: periodic calendar, overdue window, E-9 exemption, drafts,
#Concurrent-only reports, dismissals and detachment reports.
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/reports-due-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    foreach($opened in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$opened,2)}
    $app.LoadFromText(5,'ReportsDueHarness',(Join-Path $root 'tests/ReportsDueHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestReportsDue()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
