param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/group-correctness-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true);$app.LoadFromText(5,'GroupCorrectnessHarness',(Join-Path $root 'tests/GroupCorrectnessHarness.bas'));$app.DoCmd.RunCommand(126)
$result=$app.Eval('TestGroupCorrectness()');if($result -notlike 'PASS:*'){throw $result};Write-Output $result
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
