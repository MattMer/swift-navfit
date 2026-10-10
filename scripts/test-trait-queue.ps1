param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/trait-queue-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$false)
 $app.LoadFromText(5,'TraitQueueHarness',(Join-Path $root 'tests/TraitQueueHarness.bas'))
 $app.DoCmd.RunCommand(126);$result=$app.Eval('TestTraitQueue()');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
} finally {$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
