param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$SeedDemo)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
. (Join-Path (Split-Path $PSScriptRoot -Parent) 'test-contributor-context.ps1')
$hash=(Get-FileHash $BuildPath).Hash
$target=Join-Path $root ('build/tests/tools-'+[guid]::NewGuid()+'.accdb');Copy-Item $BuildPath $target
$app=New-Object -ComObject Access.Application;$primary=$null;$cleanup=@()
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
 if($SeedDemo){
  # This suite requires an approved synthetic report. Seed only its own copy;
  # SeedUXDemo rejects any populated workspace rather than merging by ID.
  $app.LoadFromText(5,'UXDemoSeed',(Join-Path $root 'tests/UXDemoSeed.bas'))
  $seedResult=$app.Eval('SeedUXDemoResult()')
  if($seedResult -notlike 'PASS:*'){throw $seedResult}
  $app.DoCmd.DeleteObject(5,'UXDemoSeed')
 }
 SelectSyntheticContributor $app
 $app.LoadFromText(5,'ReportToolsHarness',(Join-Path $root 'tests/ReportToolsHarness.bas'))
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestReportTools()');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $app.LoadFromText(5,'RevisionRetentionHarness',(Join-Path $root 'tests/RevisionRetentionHarness.bas'))
 $retention=$app.Eval('TestRevisionRetention()');Write-Output $retention
 if($retention -notlike 'PASS:*'){throw $retention}
 $coverage=$app.Eval('TestRevisionReferenceCoverage()');Write-Output $coverage
 if($coverage -notlike 'PASS:*'){throw $coverage}
} catch {$primary=$_} finally {
 try{if($app.Eval('CurrentProject.AllForms("frmEditor").IsLoaded')){[void]$app.Eval('UXMessageDiscard()')};[void]$app.Eval('UXLeaveEditor("Discard")')}catch{}
 try{$app.Quit(2)}catch{$cleanup+=$_.Exception.Message}
 try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+=$_.Exception.Message}
 if((Get-FileHash $BuildPath).Hash -ne $hash){$cleanup+='Source artifact changed'}
}
if($primary){if($cleanup){Write-Warning ($cleanup -join '; ')};throw $primary}
if($cleanup){throw ($cleanup -join '; ')}
