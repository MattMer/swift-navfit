param([Parameter(Mandatory=$true)][string]$BuildPath,[string]$SourcePath,[switch]$InstallSource,[switch]$CommitFaults)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$target=Join-Path $root ('build/tests/engine-exactness-'+[guid]::NewGuid()+'.accdb');Copy-Item $BuildPath $target
$hash=(Get-FileHash $BuildPath).Hash;$app=$null;$primary=$null;$cleanup=@()
if($CommitFaults -and -not $SourcePath){throw '-CommitFaults requires an existing synthetic -SourcePath'}
$phases=[Collections.Generic.List[string]]::new()
try{
 Write-Output ('Controller: engine outcome runner; disposable database: '+$target)
 $app=New-Object -ComObject Access.Application
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
 if($InstallSource){foreach($module in @('modOperations','modWorkspaceImport','modUX')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}}
 $app.LoadFromText(5,'EngineVerificationHarness',(Join-Path $root 'tests/EngineVerificationHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestEngineVerification()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
 $phases.Add($result)
 if($SourcePath){
  $SourcePath=[IO.Path]::GetFullPath($SourcePath);$sourceHash=(Get-FileHash $SourcePath).Hash
  [void]$app.Eval('ConfigureWorkspace("FAULT BOOTSTRAP","Tester")')
  if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$SourcePath+'")')){throw 'Preview failed'}
  [void]$app.Eval('PrepareEngineFailureUI()')
  if($app.Eval('UXApplySwiftWorkspace()')){throw 'Injected corruption accepted'}
  $status=$app.Eval('Forms!frmUXWorkspaceImport!lblStatus.Caption')
  if($status -notlike '*field mismatch in Report*'){throw ('Wrong rejection: '+$status)}
  if($app.Eval('DCount("*","Report")') -ne 0 -or $app.Eval('DLookup("CommandName","Workspace")') -ne 'FAULT BOOTSTRAP' -or $app.Eval('Forms!frmUXWorkspaceImport.Modal') -or -not $app.Eval('Forms!frmUXWorkspaceImport!cmdApply.Enabled') -or -not $app.Eval('Forms!frmUXWorkspaceImport!cmdClose.Enabled')){throw 'Failure rollback/UI restore failed'}
  if(-not $app.Eval('UXApplySwiftWorkspace()')){throw 'Valid retry failed'}
  if((Get-FileHash $SourcePath).Hash -ne $sourceHash){throw 'Source changed'}
  Write-Output 'PASS: injected mismatch rejected before commit; data/setup rollback, controls/modality restored, valid retry succeeds'
  $phases.Add('PASS: precommit mismatch rollback, UI restoration and valid retry')
  if($CommitFaults){
   $app.CloseCurrentDatabase()
   $committedTarget=$target+'.committed.accdb';Copy-Item -LiteralPath $BuildPath -Destination $committedTarget
   $app.OpenCurrentDatabase($committedTarget,$true)
   if($InstallSource){foreach($module in @('modOperations','modWorkspaceImport','modUX')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}}
   $migration=[IO.File]::ReadAllText((Join-Path $root 'src/modMigration.bas'))
   $anchor='w.CommitTrans: trans = False: committed = True'
   if(([regex]::Matches($migration,[regex]::Escape($anchor))).Count -ne 1){throw 'Expected production commit boundary was not found exactly once'}
   $migration=$migration.Replace($anchor,$anchor+': Err.Raise vbObjectError + 9601, "Synthetic post-commit failure", "SYNTHETIC committed import finalization failed"')
   $instrumented=$target+'.migration.bas';[IO.File]::WriteAllText($instrumented,$migration)
   $app.LoadFromText(5,'modMigration',$instrumented)
   $app.LoadFromText(5,'EngineVerificationHarness',(Join-Path $root 'tests/EngineVerificationHarness.bas'));$app.DoCmd.RunCommand(126)
   $result=$app.Eval('TestCommittedImportUI("'+$SourcePath+'")');Write-Output $result;if($result -notlike 'PASS:*'){throw $result};$phases.Add($result)
  }
 }
}catch{$primary=$_;throw}finally{
 if($app){
  try{$app.Quit(2)}catch{$cleanup+='Owned engine session close: '+$_.Exception.Message}
  try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+='Owned engine COM release: '+$_.Exception.Message}
 }
 if($cleanup.Count){if($primary){Write-Warning ($cleanup -join '; ')}else{throw ($cleanup -join '; ')}}
}
if((Get-FileHash $BuildPath).Hash -ne $hash){throw 'Template changed'}
if($SourcePath -and (Get-FileHash -LiteralPath $SourcePath).Hash -ne $sourceHash){throw 'Supplied source changed'}
@{Status='PASS';BuildPath=$BuildPath;BuildSHA256=$hash;SourcePath=$SourcePath;SourceSHA256=$sourceHash;InstalledSource=[bool]$InstallSource;PostCommitFault=[bool]$CommitFaults;Phases=@($phases.ToArray());Scope='Disposable copies; exact precommit verification and explicitly instrumented production commit boundary; no release acceptance'}|ConvertTo-Json -Depth 4|Set-Content -LiteralPath ($target+'.verification.json')
Write-Output ('Receipt: '+$target+'.verification.json')
