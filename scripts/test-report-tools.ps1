param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$hash=(Get-FileHash $BuildPath).Hash
$target=Join-Path $root ('build/tests/tools-'+[guid]::NewGuid()+'.accdb');Copy-Item $BuildPath $target
$app=New-Object -ComObject Access.Application
try {$app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true);$app.LoadFromText(5,'ReportToolsHarness',(Join-Path $root 'tests/ReportToolsHarness.bas'));$app.DoCmd.RunCommand(126);$result=$app.Eval('TestReportTools()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}} finally {try{if($app.Eval('CurrentProject.AllForms("frmEditor").IsLoaded')){[void]$app.Eval('UXMessageDiscard()')};[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
if((Get-FileHash $BuildPath).Hash -ne $hash){throw 'Source artifact changed'}
