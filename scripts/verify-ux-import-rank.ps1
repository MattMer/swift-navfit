param([int]$ReleaseNumber=113,[string]$ReleaseDate='2026-10-15',[int]$AssertionCount=668,[string]$EvidenceName='IMPORT_RANK_R113_VERIFICATION.json')
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$baseline=Get-Content -Raw (Join-Path $root 'planning/MEMBER_PSG_R112_VERIFICATION.json')|ConvertFrom-Json
$artifacts=@();$blankTables=@{}
foreach($variant in @('Blank','Demo')){
 $path=Join-Path $root ('output/ux/Swift_NAVFIT_UI_'+$variant+'_'+$ReleaseDate+'-r'+$ReleaseNumber+'.accdb')
 $app=New-Object -ComObject Access.Application
 try{
  $app.AutomationSecurity=1;$app.OpenCurrentDatabase($path,$true)
  #AutoExec may complete just after OpenCurrentDatabase returns in automation.
  for($startupWait=0;$startupWait -lt 30 -and $app.Forms.Count -eq 0;$startupWait++){Start-Sleep -Milliseconds 100}
  #Runtime forms with class modules can expose their code interface to COM;
  #AllForms supplies stable AccessObject names without that dispatch ambiguity.
  $forms=@($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name});$expected=if($variant -eq 'Blank'){'frmSetup'}else{'frmHome'}
  if($forms -notcontains $expected){throw ('Startup form missing: '+$variant)}
  $count=[int]$app.Eval('DCount("*","Report")');if($count -ne $(if($variant -eq 'Blank'){0}else{9})){throw 'Unexpected report count'}
  if($variant -eq 'Blank'){
   foreach($table in $baseline.BlankTables.PSObject.Properties.Name){$rows=[int]$app.Eval('DCount("*","'+$table+'")');$blankTables[$table]=$rows;if($rows -ne 0){throw ('Nonempty blank table: '+$table)}}
  }
  foreach($name in $forms){$app.DoCmd.Close(2,$name,2)}
  foreach($screen in @('frmPDFImport','frmUXNativeGroups','frmImportReview','frmNAVFITExchange','frmUXGroupReport','frmUXDirectoryBatch')){
   $app.DoCmd.OpenForm($screen,1);$f=$app.Forms.Item($screen);$controls=@($f.Controls|ForEach-Object {$_.Name})
   foreach($name in @('cboCentury','cboNAVFITCentury','cboReturnCentury')){if($controls -contains $name){throw 'Century selector retained'}}
   if($screen -in @('frmUXGroupReport','frmUXDirectoryBatch') -and $controls -notcontains 'cboRankFilter'){throw 'Missing rank filter'}
   $app.DoCmd.Close(2,$screen,2)
  }
  $app.DoCmd.RunCommand(126)
 }finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 $f=$null
 for($attempt=0;$attempt -lt 30;$attempt++){try{$hash=(Get-FileHash -LiteralPath $path -ErrorAction Stop).Hash;break}catch{if($attempt -eq 29){throw};Start-Sleep -Milliseconds 200}}
 $artifacts+=@{Path=$path;StartupForms=$forms;Reports=$count;SHA256=$hash;Bytes=(Get-Item -LiteralPath $path).Length}
}
@{Candidate=('r'+$ReleaseNumber);Date=$ReleaseDate;AppVersion='0.25.0-dev';Schema=25;Assertions=$AssertionCount;Artifacts=$artifacts;BlankTables=$blankTables;ManualDisplayVerification=$false}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $root ('planning/'+$EvidenceName))
Write-Output ('PASS: both r'+$ReleaseNumber+' builds compile/start; century selectors absent; rank filters present; blank tables empty; nine synthetic demo reports')
