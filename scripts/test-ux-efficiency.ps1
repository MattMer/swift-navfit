param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/efficiency-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$false)
 $app.LoadFromText(5,'EfficiencyHarness',(Join-Path $root 'tests/EfficiencyHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestEfficiency()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
 if(-not $app.Eval('PrepareEventTimer()')){throw 'Timer preparation failed'}
 Start-Sleep -Seconds 31
 $result=$app.Eval('FinishEventTimer()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
