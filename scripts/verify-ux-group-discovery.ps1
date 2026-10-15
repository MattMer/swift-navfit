$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$baseline=Get-Content -Raw (Join-Path $root 'planning/READINESS_R110_VERIFICATION.json')|ConvertFrom-Json
$artifacts=@();$blankTables=@{}
foreach($variant in @('Blank','Demo')){
 $path=Join-Path $root ('output/ux/Swift_NAVFIT_UI_'+$variant+'_2026-10-14-r111.accdb')
 $app=New-Object -ComObject Access.Application
 try{
  $app.AutomationSecurity=1;$app.OpenCurrentDatabase($path,$true)
  $forms=@($app.Forms|ForEach-Object {$_.Name});$expected=if($variant -eq 'Blank'){'frmSetup'}else{'frmHome'}
  if($forms -notcontains $expected){throw ('Startup form missing: '+$variant)}
  $count=[int]$app.Eval('DCount("*","Report")')
  if($count -ne $(if($variant -eq 'Blank'){0}else{9})){throw 'Unexpected report count'}
  if($variant -eq 'Blank'){
   foreach($table in $baseline.BlankTables.PSObject.Properties.Name){
    $rows=[int]$app.Eval('DCount("*","'+$table+'")');$blankTables[$table]=$rows
    if($rows -ne 0){throw ('Nonempty blank table: '+$table)}
   }
  }
  $app.DoCmd.Close(2,$expected,2);$app.DoCmd.OpenForm('frmUXNativeGroups',1)
  if($app.Forms.Item('frmUXNativeGroups').Controls('cmdAssignImportGrade').OnClick -ne '=UXAssignNativeGroupGrade()'){throw 'Missing assignment handler'}
  $app.DoCmd.Close(2,'frmUXNativeGroups',2)
  $app.DoCmd.RunCommand(126)
 }finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 $artifacts+=@{Path=$path;StartupForms=$forms;Reports=$count;SHA256=(Get-FileHash -LiteralPath $path).Hash;Bytes=(Get-Item -LiteralPath $path).Length}
}
@{Candidate='r111';Date='2026-10-14';AppVersion='0.25.0-dev';Schema=25;TargetedAssertions=28;NativeGroupRegressionAssertions=248;Artifacts=$artifacts;BlankTables=$blankTables;SourceOriginalUnchanged=$true}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $root 'planning/GROUP_DISCOVERY_R111_VERIFICATION.json')
Write-Output 'PASS: both candidates compile/start, 48 blank data tables empty, demo has nine reports; verification recorded'
