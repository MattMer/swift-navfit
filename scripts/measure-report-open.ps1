param([Parameter(Mandatory=$true)][string]$BuildPath,[Parameter(Mandatory=$true)][string]$BaselinePath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent;$results=@()
foreach($entry in @(@{Name='r114';Path=$BaselinePath},@{Name='r115';Path=$BuildPath})){
 $scratch=Join-Path $root ('build/tests/report-timing-'+$entry.Name+'-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $entry.Path -Destination $scratch
 $app=New-Object -ComObject Access.Application
 try{
  $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
  foreach($form in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$form,2)}
  [void]$app.Eval('ConfigureWorkspace("SYNTHETIC TIMING","Tester")')
  $id=$app.Eval('CommitDraft(0,"EVAL","Timing","SYNTHETIC, MEMBER",#1/1/2026#,#9/30/2026#,"Synthetic narrative")')
  $opens=@();$validation=@()
  for($i=0;$i -lt 4;$i++){
   if($app.CurrentProject.AllForms.Item('frmEditor').IsLoaded){throw 'Editor was not closed'}
   $watch=[Diagnostics.Stopwatch]::StartNew();$app.DoCmd.OpenForm('frmEditor',0,[Type]::Missing,[Type]::Missing,[Type]::Missing,[Type]::Missing,[string]$id);$watch.Stop();$opens+=$watch.Elapsed.TotalMilliseconds
   $watch.Restart();[void]$app.Eval('UXLiveTimer()');$watch.Stop();$validation+=$watch.Elapsed.TotalMilliseconds
   if(-not $app.Eval('UXLeaveEditor("Discard")')){throw 'Editor could not close'}
   if($app.CurrentProject.AllForms.Item('frmEditor').IsLoaded){throw 'Editor still loaded after close'}
  }
  $results+=@{Build=$entry.Name;OpenMilliseconds=$opens;InitialValidationMilliseconds=$validation;ReopenedEachTime=$true}
 }finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
$results|ConvertTo-Json -Depth 5|Set-Content (Join-Path $root 'build/r115-report-open-timing.json')
$results|ConvertTo-Json -Depth 5
