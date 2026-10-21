param([Parameter(Mandatory=$true)][string]$BuildPath,[string]$SourcePath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$target=Join-Path $root ('build/tests/engine-exactness-'+[guid]::NewGuid()+'.accdb');Copy-Item $BuildPath $target
$hash=(Get-FileHash $BuildPath).Hash;$app=New-Object -ComObject Access.Application
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
 $app.LoadFromText(5,'EngineVerificationHarness',(Join-Path $root 'tests/EngineVerificationHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestEngineVerification()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
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
 }
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
if((Get-FileHash $BuildPath).Hash -ne $hash){throw 'Template changed'}
