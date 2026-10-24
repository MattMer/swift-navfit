param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$EntryLabelsOnly,[switch]$InstallSource,[switch]$EditorReset)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/compat-pages-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
$primary=$null;$cleanup=@()
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 if($InstallSource){
  $app.DoCmd.OpenForm('frmSetup',1);$f=$app.Forms.Item('frmSetup')
  foreach($spec in @(@('lblCommand','txtCommand',1950,'Command name'),@('lblUser','txtProfile',2750,'Contributor name'))){
   $app.DeleteControl($f.Name,$spec[0]);$label=$app.CreateControl($f.Name,100,0,$spec[1],'',400,$spec[2],2800,350);$label.Name=$spec[0];$label.Caption=$spec[3];$label.FontName='Segoe UI';$label.FontSize=10
  }
  $app.DoCmd.Close(2,$f.Name,1)
  . (Join-Path $PSScriptRoot 'apply-contributor-startup.ps1')
  . (Join-Path $PSScriptRoot 'apply-pdf-output.ps1')
 }
 $app.LoadFromText(5,'CompatPagesHarness',(Join-Path $root 'tests/CompatPagesHarness.bas'));$app.DoCmd.RunCommand(126)
 if(-not $EntryLabelsOnly){
 $result=$app.Eval('TestCompatPages()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
 }
 $result=$app.Eval('TestEntryLabels()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
 if($EditorReset){
  $result=$app.Eval('PrepareEditorResetFeedback()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
  $app.VBE.CommandBars.FindControl(1,228).Execute()
  $result=$app.Eval('VerifyEditorResetFeedback()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
 }
}catch{$primary=$_;throw}finally{
 try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{}
 try{$app.Quit(2)}catch{$cleanup+=$_.Exception.Message}
 try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+=$_.Exception.Message}
 if($cleanup){if($primary){Write-Warning ($cleanup -join '; ')}else{throw ($cleanup -join '; ')}}
}
