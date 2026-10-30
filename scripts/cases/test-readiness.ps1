param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/readiness-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'ReadinessHarness',(Join-Path $root 'tests/ReadinessHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestReadiness()');Write-Output $result
 @{Result=$result;Scratch=$scratch;SHA256=(Get-FileHash $BuildPath).Hash}|ConvertTo-Json|Set-Content (Join-Path $root 'build/tests/readiness-results.json')
 if($result -notlike 'PASS:*'){throw $result}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
