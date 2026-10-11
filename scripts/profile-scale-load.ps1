$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$fixture=Get-Content (Join-Path $root 'planning/SCALE_R102_FIXTURE.json') -Raw|ConvertFrom-Json
$scratch=Join-Path $root ('build/tests/scale-profile-'+[guid]::NewGuid()+'.accdb');Copy-Item -LiteralPath $fixture.Database -Destination $scratch
$app=New-Object -ComObject Access.Application;$rows=@()
function Timed([string]$expression){$watch=[Diagnostics.Stopwatch]::StartNew();$value=$app.Eval($expression);$watch.Stop();if(-not $value){throw $expression};[double]$watch.Elapsed.TotalMilliseconds}
function Median($values){$sorted=@($values|Sort-Object);[Math]::Round($sorted[[int][Math]::Floor($sorted.Count/2)],2)}
try{
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$false);$app.DoCmd.OpenForm('frmUXGroups')
 foreach($g in @($fixture.Groups|Where-Object {$_.Kind -eq 'EVAL'})){
  $open=Timed ('ScaleOpen('+$g.GroupID+')');$reloads=@();$averages=@()
  for($i=0;$i -lt 3;$i++){$reloads+=Timed 'UXGridLoad()'}
  for($i=0;$i -lt 5;$i++){$averages+=Timed 'UXGridRefreshAverages()'}
  $row=@{Size=$g.Size;OpenMs=$open;CacheReloadMedianMs=Median $reloads;AverageRefreshMedianMs=Median $averages;CacheReloadRawMs=$reloads;AverageRefreshRawMs=$averages};$rows+=$row
  Write-Output ('Profile EVAL '+$g.Size+': cache reload '+$row.CacheReloadMedianMs+' ms; averages '+$row.AverageRefreshMedianMs+' ms')
  if(-not $app.Eval('UXCloseComparison()')){throw 'Close failed'}
 }
 @{Date='2026-10-10';Basis='Separate scratch copy/session; EVAL groups only; 3 cache reloads and 5 unchanged average refreshes per size; COM stopwatch includes overhead';Rows=$rows}|ConvertTo-Json -Depth 6|Set-Content (Join-Path $root 'planning/SCALE_R102_LOAD_PROFILE.json') -Encoding UTF8
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
