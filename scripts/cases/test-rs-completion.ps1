param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/rs-completion-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true);$app.LoadFromText(5,'RSCompletionHarness',(Join-Path $root 'tests/RSCompletionHarness.bas'));$app.DoCmd.RunCommand(126)
$expression='TestRSCompletion()'
$result=$app.Eval($expression);Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
}finally{try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
