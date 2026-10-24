param([Parameter(Mandatory=$true)][string]$TemplatePath,[Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath){throw 'Refusing overwrite'}
$TemplatePath=[IO.Path]::GetFullPath($TemplatePath)
$sourceHash=(Get-FileHash -LiteralPath $TemplatePath).Hash
# Reject populated templates through read-only DAO before creating a destination.
& powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'verify-build-artifact.ps1') -BuildPath $TemplatePath -RequireBlank
if($LASTEXITCODE -ne 0){throw 'Blank template verification failed; demo not created'}
[void](New-Item -ItemType Directory -Force -Path (Split-Path $OutputPath -Parent))
Copy-Item -LiteralPath $TemplatePath -Destination $OutputPath
$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
 foreach($module in @('UXDemoSeed','DirectoryDemoSeed')){$app.LoadFromText(5,$module,(Join-Path $root ('tests/'+$module+'.bas')))}
 foreach($expression in @('SeedUXDemoResult()','SeedDirectoryDemoHistory()')){
  $result=$app.Eval($expression);Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
 }
 [void]$app.Eval('MigrateLegacyReportNotes()');[void]$app.Eval('MigrateReportMessages()')
 if(-not $app.Eval('EnsureStarterSnippets()')){throw 'Snippet seeding failed'}
 foreach($module in @('UXDemoSeed','DirectoryDemoSeed')){$app.DoCmd.DeleteObject(5,$module)}
 # Seed fixture simulates a crash; restore the saved native close handler afterwards.
 $app.DoCmd.OpenForm('frmEditor',1);$app.Forms.Item('frmEditor').OnUnload='[Event Procedure]';$app.DoCmd.Close(2,'frmEditor',1)
 $app.LoadFromText(4,'AutoExec',(Join-Path $root 'src/AutoExec.txt'))
 if(-not @($app.CurrentProject.AllMacros|ForEach-Object {$_.Name}).Contains('AutoExec')){throw 'Missing startup macro'}
 $app.DoCmd.RunCommand(126);$app.CloseCurrentDatabase()
} finally {$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
if((Get-FileHash -LiteralPath $TemplatePath).Hash -ne $sourceHash){throw 'Demo build changed its source template'}
foreach($suite in @('test-ux-demo','test-preview-launch')){
 & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot ($suite+'.ps1')) -BuildPath $OutputPath
 if($LASTEXITCODE -ne 0){throw ($suite+' failed; demo is not verified')}
}
if((Get-FileHash -LiteralPath $TemplatePath).Hash -ne $sourceHash){throw 'Demo checks changed the source template'}
Write-Output ('PASS: compiled synthetic demo, populated workflow and startup; source unchanged: '+$OutputPath)
