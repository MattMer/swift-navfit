param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource,[switch]$ReusePresentation)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$folder=Join-Path $root ('build/tests/pdf-output-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
$scratch=Join-Path $folder 'fixture.accdb';Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
Write-Output ('Controller: focused PDF runner; database: '+$scratch)
try {
 $app.AutomationSecurity=1;$app.Visible=$true;$app.OpenCurrentDatabase($scratch,$true)
 if($InstallSource){
  . (Join-Path $PSScriptRoot 'apply-pdf-print-catalogue.ps1')
  foreach($module in @('modPDFOutput','modPDFOutputUI','modUXTraitGrid','modReportPages','modUXGroupPages')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}
  if($ReusePresentation){. (Join-Path $PSScriptRoot 'apply-pdf-output.ps1') -RefreshBackgroundsOnly}else{. (Join-Path $PSScriptRoot 'apply-pdf-output.ps1')}
 }
 $app.LoadFromText(5,'PDFOutputHarness',(Join-Path $root 'tests/PDFOutputHarness.bas'))
 Write-Output 'Compile focused PDF source'
 $app.DoCmd.RunCommand(126)
 Write-Output 'Run PDF behavior checks'
 $result=$app.Eval('TestPDFOutput("'+$folder.Replace('"','""')+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $result=$app.Eval('TestPDFDraftsAndKinds("'+$folder.Replace('"','""')+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
 $result=$app.Eval('TestPDFBoundaries("'+$folder.Replace('"','""')+'")');Write-Output $result
 if($result -notlike 'PASS:*'){throw $result}
} catch {
 $primary=$_
 try{$pane=$app.VBE.ActiveCodePane;$cm=$pane.CodeModule;Write-Output ('Active VBA module: '+$cm.Name)}catch{}
 throw $primary
} finally {
 try{$app.Quit(2)}catch{Write-Warning ('Owned Access cleanup: '+$_.Exception.Message)}
 try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{Write-Warning ('Owned COM release: '+$_.Exception.Message)}
}
Write-Output ('Fixture and PDFs: '+$folder)
