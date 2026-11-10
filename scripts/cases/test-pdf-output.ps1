param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource,[switch]$ReusePresentation,[switch]$MetricsOnly,[switch]$ChiefProofOnly,[switch]$ChiefFieldsOnly,[switch]$ChiefCasesOnly)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$folder=Join-Path $root ('build/tests/pdf-output-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
$scratch=Join-Path $folder 'fixture.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$results=@()
$primary=$null;$cleanup=@()
$checkerHash=(Get-FileHash -LiteralPath $PSCommandPath).Hash
$harnessHash=(Get-FileHash -LiteralPath (Join-Path $root 'tests/PDFOutputHarness.bas')).Hash
$app=New-Object -ComObject Access.Application
Write-Output ('Controller: focused PDF runner; database: '+$scratch)
try {
 $app.AutomationSecurity=1;$app.Visible=$true;$app.OpenCurrentDatabase($scratch,$true)
 if($InstallSource){
  # Focused-copy preparation only; production upgrade uses fresh-build/import.
  $definitions=Get-Content (Join-Path $root 'src/field-definitions.json') -Raw|ConvertFrom-Json
  $db=$app.CurrentDb()
  foreach($field in $definitions.fields|Where-Object {$_.definitionVersion -ge 2}){
   if(@($db.TableDefs.Item('Report').Fields|ForEach-Object {$_.Name}) -notcontains $field.key){$db.Execute(('ALTER TABLE Report ADD COLUMN ['+$field.key+'] '+$field.storageType),128)}
   if(@($db.TableDefs.Item('RecoveryDraft').Fields|ForEach-Object {$_.Name}) -notcontains $field.recoveryKey){$db.Execute(('ALTER TABLE RecoveryDraft ADD COLUMN ['+$field.recoveryKey+'] LONGTEXT'),128)}
  }
  $db.Execute("UPDATE AppMetadata SET SchemaVersion=33,AppVersion='0.36.0-dev'",128);$db.Close();$db=$null
  . (Join-Path (Split-Path $PSScriptRoot -Parent) 'apply-ux-field-catalogue.ps1')
  . (Join-Path (Split-Path $PSScriptRoot -Parent) 'apply-ux-surface-catalogue.ps1')
  #Forms are source: reload them from src/forms.
  . (Join-Path (Split-Path $PSScriptRoot -Parent) 'install-form-text.ps1');[void](InstallSwiftFormText)
  . (Join-Path (Split-Path $PSScriptRoot -Parent) 'apply-pdf-print-catalogue.ps1')
  foreach($module in @('modSwift','modMigration','modRevisionCompare','modNAVFITAdapter','modNAVFITExchange','modUXNAVFIT','modPDFOutput','modPDFOutputUI','modUXTraitGrid','modReportPages','modUXGroupPages')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}
  if($ReusePresentation){. (Join-Path (Split-Path $PSScriptRoot -Parent) 'resource-pdf-report.ps1') -RefreshBackgroundsOnly}else{. (Join-Path (Split-Path $PSScriptRoot -Parent) 'resource-pdf-report.ps1')}
 }
 $app.LoadFromText(5,'PDFOutputHarness',(Join-Path $root 'tests/PDFOutputHarness.bas'))
 Write-Output 'Compile focused PDF source'
 $app.DoCmd.RunCommand(126)
 $result=$app.Eval('TestPDFMetrics()');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $results+=$result
 if($MetricsOnly){return}
 if($ChiefCasesOnly){
  $result=$app.Eval('TestChiefPrintCases("'+$folder.Replace('"','""')+'")');Write-Output $result
  if($result -notlike 'PASS:*'){throw $result}
  $results+=$result
  return
 }
 if($ChiefFieldsOnly){
  $result=$app.Eval('TestChiefComments("'+$folder.Replace('"','""')+'","'+(Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb')+'")');Write-Output $result
  if($result -notlike 'PASS:*'){throw $result}
  $results+=$result
  return
 }
 if($ChiefProofOnly){
  $result=$app.Eval('TestChiefRenderingProof("'+$folder.Replace('"','""')+'")');Write-Output $result
  if($result -notlike 'PASS:*'){throw $result}
  $results+=$result
  return
 }
 Write-Output 'Run PDF behavior checks'
 $result=$app.Eval('TestPDFOutput("'+$folder.Replace('"','""')+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $results+=$result
 $result=$app.Eval('TestPDFDraftsAndKinds("'+$folder.Replace('"','""')+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $results+=$result
 $result=$app.Eval('TestPDFLayoutPrinter("'+$folder.Replace('"','""')+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $results+=$result
 $result=$app.Eval('TestPDFBoundaries("'+$folder.Replace('"','""')+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $results+=$result
 $result=$app.Eval('TestChiefComments("'+$folder.Replace('"','""')+'","'+(Join-Path $root 'FORMS_REFERENCES_TEMPLATES/NAVFIT_BLANK.accdb')+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $results+=$result
 $result=$app.Eval('TestChiefRenderingProof("'+$folder.Replace('"','""')+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $results+=$result
 $result=$app.Eval('TestChiefPrintCases("'+$folder.Replace('"','""')+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $results+=$result
} catch {
 $primary=$_
 try{$pane=$app.VBE.ActiveCodePane;$cm=$pane.CodeModule;Write-Output ('Active VBA module: '+$cm.Name)}catch{}
 throw $primary
} finally {
 try{$app.Quit(2)}catch{$cleanup+='Owned Access cleanup: '+$_.Exception.Message}
 try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+='Owned COM release: '+$_.Exception.Message}
 [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
 if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $buildHash){$cleanup+='Focused PDF runner changed the supplied build'}
 if((Get-FileHash -LiteralPath $PSCommandPath).Hash -ne $checkerHash -or (Get-FileHash -LiteralPath (Join-Path $root 'tests/PDFOutputHarness.bas')).Hash -ne $harnessHash){$cleanup+='Focused PDF checker/harness changed during the run'}
 @{Status=$(if($primary -or $cleanup){'FAIL'}else{'PASS'});Results=$results;BuildSHA256=$buildHash;InstalledSource=[bool]$InstallSource;Fixture=$scratch;CheckerSHA256=$checkerHash;HarnessSHA256=$harnessHash}|ConvertTo-Json|Set-Content (Join-Path $folder 'verification.json')
 if($cleanup){if($primary){Write-Warning ($cleanup -join '; ')}else{throw ($cleanup -join '; ')}}
}
Write-Output ('Fixture and PDFs: '+$folder)
