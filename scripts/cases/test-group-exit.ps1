param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$TaskOrders,[switch]$InstallSource)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/navigation-context-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
if($InstallSource){foreach($module in @('modUXResourceTools','modUXNAVFIT')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}}
$app.LoadFromText(5,'GroupExitHarness',(Join-Path $root 'tests/GroupExitHarness.bas'));$app.DoCmd.RunCommand(126)
$result=$app.Eval('TestGroupExit()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
$result=$app.Eval('TestUnifiedMembers()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
if($TaskOrders){$result=$app.Eval('TestGroupTaskOrders()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}}
}finally{try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
Write-Output ('Fixture: '+$scratch)
