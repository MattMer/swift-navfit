$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$baseline=Get-Content -Raw (Join-Path $root 'planning/GROUP_DISCOVERY_R111_VERIFICATION.json')|ConvertFrom-Json
$artifacts=@();$blankTables=@{}
foreach($variant in @('Blank','Demo')){
 $path=Join-Path $root ('output/ux/Swift_NAVFIT_UI_'+$variant+'_2026-10-15-r112.accdb')
 $app=New-Object -ComObject Access.Application
 try{
  $app.AutomationSecurity=1;$app.OpenCurrentDatabase($path,$true)
  $forms=@($app.Forms|ForEach-Object {$_.Name});$expected=if($variant -eq 'Blank'){'frmSetup'}else{'frmHome'}
  if($forms -notcontains $expected){throw ('Startup form missing: '+$variant)}
  $count=[int]$app.Eval('DCount("*","Report")');if($count -ne $(if($variant -eq 'Blank'){0}else{9})){throw 'Unexpected report count'}
  if($variant -eq 'Blank'){
   foreach($table in $baseline.BlankTables.PSObject.Properties.Name){$rows=[int]$app.Eval('DCount("*","'+$table+'")');$blankTables[$table]=$rows;if($rows -ne 0){throw ('Nonempty blank table: '+$table)}}
  }
  foreach($name in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
  $app.DoCmd.OpenForm('frmUXGroupRSCA',1);$f=$app.Forms.Item('frmUXGroupRSCA')
  if($f.Controls('cboBaseline').AfterUpdate -ne '=RSGroupBaselineChanged()' -or $f.Controls('chkExcludes').AfterUpdate -ne '=RSSaveGroupCalculation()'){throw 'Autosave handlers missing'}
  if(@($f.Controls|ForEach-Object {$_.Name}) -contains 'cmdSave'){throw 'Redundant save button retained'}
  $app.DoCmd.Close(2,'frmUXGroupRSCA',2);$app.DoCmd.RunCommand(126)
  $export=Join-Path $root ('build/tests/r112-'+$variant+'-directory-source.bas');$app.SaveAsText(5,'modCommandDirectory',$export)
  if((Get-Content -Raw $export) -notmatch 'Public Sub RefreshDirectoryFromReport'){throw 'Final enrichment service not embedded'}
 }finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 $f=$null
 for($attempt=0;$attempt -lt 30;$attempt++){
  try{$hash=(Get-FileHash -LiteralPath $path -ErrorAction Stop).Hash;break}catch{if($attempt -eq 29){throw};Start-Sleep -Milliseconds 200}
 }
 $artifacts+=@{Path=$path;StartupForms=$forms;Reports=$count;SHA256=$hash;Bytes=(Get-Item -LiteralPath $path).Length}
}
@{Candidate='r112';Date='2026-10-15';AppVersion='0.25.0-dev';Schema=25;Assertions=721;Artifacts=$artifacts;BlankTables=$blankTables;ManualDisplayVerification=$false}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $root 'planning/MEMBER_PSG_R112_VERIFICATION.json')
Write-Output 'PASS: both r112 candidates compile/start, autosave handlers embedded, blank tables empty, synthetic demo preserved'
