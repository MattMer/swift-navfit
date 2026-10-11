param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/performance-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$false)
 $app.LoadFromText(5,'PerformanceHarness',(Join-Path $root 'tests/PerformanceHarness.bas'))
 $app.DoCmd.RunCommand(126);$result=$app.Eval('TestPerformance()');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
} finally {$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
