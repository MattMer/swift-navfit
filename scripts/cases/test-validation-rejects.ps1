param([Parameter(Mandatory=$true)][string]$BuildPath)
# Checks drawn from PERS-32 reject reasons (tests/ValidationRejectsHarness.bas).
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/validation-rejects-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 $app.LoadFromText(5,'ValidationRejectsHarness',(Join-Path $root 'tests/ValidationRejectsHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=[string]$app.Eval('TestValidationRejects()');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
