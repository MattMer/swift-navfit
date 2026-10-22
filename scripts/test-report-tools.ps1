param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$SeedDemo)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'test-contributor-context.ps1')
$hash=(Get-FileHash $BuildPath).Hash
$target=Join-Path $root ('build/tests/tools-'+[guid]::NewGuid()+'.accdb');Copy-Item $BuildPath $target
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
 if($SeedDemo){
  # This suite requires an approved synthetic report. Seed only its own copy;
  # SeedUXDemo rejects any populated workspace rather than merging by ID.
  $app.LoadFromText(5,'UXDemoSeed',(Join-Path $root 'tests/UXDemoSeed.bas'))
  $seedResult=$app.Eval('SeedUXDemo()')
  if($seedResult -notlike 'PASS:*'){throw $seedResult}
  $app.DoCmd.DeleteObject(5,'UXDemoSeed')
 }
 SelectSyntheticContributor $app
 $app.LoadFromText(5,'ReportToolsHarness',(Join-Path $root 'tests/ReportToolsHarness.bas'))
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestReportTools()');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
} finally {try{if($app.Eval('CurrentProject.AllForms("frmEditor").IsLoaded')){[void]$app.Eval('UXMessageDiscard()')};[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
if((Get-FileHash $BuildPath).Hash -ne $hash){throw 'Source artifact changed'}
