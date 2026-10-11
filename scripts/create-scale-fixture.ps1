param([string]$TemplatePath='output/ux/Swift_NAVFIT_UI_Blank_2026-10-10-r102.accdb',[string]$OutputPath='output/ux/Swift_NAVFIT_Scale_Test_2026-10-10-r102.accdb')
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent;$OutputPath=[IO.Path]::GetFullPath((Join-Path $root $OutputPath))
if(Test-Path -LiteralPath $OutputPath){throw 'Refusing to overwrite scale fixture'}
Copy-Item -LiteralPath (Join-Path $root $TemplatePath) -Destination $OutputPath
$app=New-Object -ComObject Access.Application;$groups=@()
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($OutputPath,$true)
 foreach($form in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$form,2)}
 $app.LoadFromText(5,'ScaleHarness',(Join-Path $root 'tests/ScaleHarness.bas'));$app.DoCmd.RunCommand(126)
 if(-not $app.Eval('ConfigureWorkspace("SYNTHETIC SCALE TEST - NO REAL PERSONNEL","Scale Tester")')){throw 'Setup failed'}
 foreach($kind in @('EVAL','CHIEFEVAL','FITREP')){foreach($size in @(25,50,100)){
  $watch=[Diagnostics.Stopwatch]::StartNew();$id=$app.Eval('SeedScaleGroup('+ $size +',"'+ $kind +'")');$watch.Stop()
  $actual=$app.Eval('DCount("*","GroupMember","GroupID='+$id+'")');if($actual -ne $size){throw 'Fixture count mismatch'}
  $groups+=@{Kind=$kind;Size=$size;GroupID=$id;SeedMilliseconds=$watch.Elapsed.TotalMilliseconds}
  Write-Output ('Created '+$kind+' group '+$id+': '+$actual+' synthetic reports')
 }}
 $reportCount=$app.Eval('DCount("*","Report")');if($reportCount -ne 525){throw 'Expected 525 reports'}
 @{Database=$OutputPath;ReportCount=$reportCount;Groups=$groups;SyntheticOnly=$true}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $root 'planning/SCALE_R102_FIXTURE.json') -Encoding UTF8
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
Write-Output ('Scale fixture: '+$OutputPath)
