param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/group-correctness-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true);$app.LoadFromText(5,'RecursiveDefaultsHarness',(Join-Path $root 'tests/RecursiveDefaultsHarness.bas'));$app.LoadFromText(5,'UXGlobalDefaultsHarness',(Join-Path $root 'tests/UXGlobalDefaultsHarness.bas'));$app.DoCmd.RunCommand(126)
$result=$app.Eval('TestUXGlobalDefaults()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
}finally{try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
