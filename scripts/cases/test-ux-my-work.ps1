param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/my-work-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
if($InstallSource){foreach($module in @('modUX','modUXMyWork')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}}
$app.LoadFromText(5,'UXMyWorkHarness',(Join-Path $root 'tests/UXMyWorkHarness.bas'));$app.DoCmd.RunCommand(126)
$result=$app.Eval('TestUXMyWork()');if(-not $result.StartsWith('PASS:')){throw $result};Write-Output $result
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
