param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$EntryLabelsOnly,[switch]$InstallSource,[switch]$EditorReset)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/compat-pages-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
$primary=$null;$cleanup=@()
try{$app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
 if($InstallSource){
  #Forms are source: reload them from src/forms, with the report they open.
  . (Join-Path $PSScriptRoot 'install-form-text.ps1');[void](InstallSwiftFormText)
  . (Join-Path $PSScriptRoot 'resource-pdf-report.ps1')
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
