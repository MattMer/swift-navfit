param([Parameter(Mandatory=$true)][string]$BuildPath,[Parameter(Mandatory=$true)][string]$BaselinePath,[switch]$InstallSource,[int]$Size=500,[int]$Trials=5,[string]$DataSourcePath)
$ErrorActionPreference='Stop';$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
if($Size -lt 10 -or $Size -gt 1000 -or $Trials -lt 3){throw 'Use 10-1000 rows and at least three trials'}
$folder=Join-Path $root ('build/tests/batched-reads-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
$originals=@{};foreach($path in @($BuildPath,$BaselinePath)){$originals[[IO.Path]::GetFullPath($path)]=(Get-FileHash -LiteralPath $path).Hash}
if($DataSourcePath){$originals[[IO.Path]::GetFullPath($DataSourcePath)]=(Get-FileHash -LiteralPath $DataSourcePath).Hash}
$samples=[Collections.Generic.List[object]]::new();$expected=@{};$app=$null;$cleanupFailure=''
$baseline=Join-Path $folder 'baseline.accdb';$candidate=Join-Path $folder 'candidate.accdb'
Copy-Item -LiteralPath $BaselinePath -Destination $baseline;Copy-Item -LiteralPath $BuildPath -Destination $candidate
try{
 foreach($mode in @('Baseline','Candidate')){
  $path=if($mode -eq 'Baseline'){$baseline}else{$candidate}
  Write-Output ('Controller: batched-read runner; database: '+$path)
  $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($path,$true)
  if($mode -eq 'Candidate' -and $InstallSource){
   foreach($module in @('modReportState','modRSDirectory','modWorkspace','modBulkFolders','modMemberHistory','modUXReadRows','modUXDirectory','modSwift')){$app.LoadFromText(5,$module,(Join-Path $root ('src/'+$module+'.bas')))}
  }
  $app.LoadFromText(5,'ReadBenchmarkHarness',(Join-Path $root 'tests/ReadBenchmarkHarness.bas'))
  if($mode -eq 'Candidate'){$app.LoadFromText(5,'BatchedReadsHarness',(Join-Path $root 'tests/BatchedReadsHarness.bas'))}
  $app.DoCmd.RunCommand(126)
  if($mode -eq 'Baseline'){
   if($DataSourcePath){
    $dataPath=[IO.Path]::GetFullPath($DataSourcePath)
    $result=$app.Eval('MigrateWorkspace("'+$dataPath+'",False,False,"'+$originals[$dataPath]+'")');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
    [void]$app.Eval('SwitchContributor(1)')
    if($app.Eval('DLookup("CommandName","Workspace")') -cne 'SYNTHETIC BATCHED READS' -or $app.Eval('DCount("*","Report")') -ne $Size -or $app.Eval('DCount("*","PriorEvaluation")') -ne $Size){throw 'Data source must be the unchanged synthetic benchmark fixture of the requested size'}
   }else{$result=$app.Eval('SeedReadBenchmark('+$Size+')');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}}
  }else{
   $result=$app.Eval('MigrateWorkspace("'+$baseline+'",False,False,"'+$baselineHash+'")');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
   [void]$app.Eval('SwitchContributor(1)')
  }
  [void]$app.Eval('PrepareReadViews()')
  $before=$app.Eval('ReadHistoryCounts()')
  foreach($case in @('Folder stamp','Directory history','Group RSCA')){
   $value=[string]$app.Eval('ReadBenchmarkCase("'+$case+'")')
   if($mode -eq 'Baseline'){$expected[$case]=$value}elseif($value -cne $expected[$case]){throw ('Read output differs: '+$case)}
   for($trial=1;$trial -le $Trials;$trial++){
    $timer=[Diagnostics.Stopwatch]::StartNew();$value=[string]$app.Eval('ReadBenchmarkCase("'+$case+'")');$timer.Stop()
    if($value -cne $expected[$case]){throw ('Unstable read output: '+$case)}
    $samples.Add([pscustomobject]@{Mode=$mode;Case=$case;Rows=$Size;Trial=$trial;Milliseconds=$timer.Elapsed.TotalMilliseconds})
   }
   Write-Output ('PASS: '+$mode+' '+$case+' output/order, '+$Trials+' samples')
  }
  if($app.Eval('ReadHistoryCounts()') -cne $before){throw 'Display/calculation reads wrote data or history'}
  # Save completed measurements before guard/fault tests; a test failure must not
  # lose timing evidence or turn partial coverage into a PASS receipt.
  $samples|Export-Csv -LiteralPath (Join-Path $folder 'timings.csv') -NoTypeInformation
  if($mode -eq 'Candidate'){
   $app.CloseCurrentDatabase()
   $guard=Join-Path $folder 'guards.accdb';Copy-Item -LiteralPath $candidate -Destination $guard
   $app.OpenCurrentDatabase($guard,$true);[void]$app.Eval('SwitchContributor(1)')
   $result=$app.Eval('TestBatchedReads()');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
  }
  $app.CloseCurrentDatabase();$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
  if($mode -eq 'Baseline'){$baselineHash=(Get-FileHash -LiteralPath $baseline).Hash}
 }
}catch{$primary=$_;throw $primary}
finally{if($app){try{$app.Quit(2)}catch{$cleanupFailure=$_.Exception.Message;Write-Warning $cleanupFailure};try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{Write-Warning $_.Exception.Message}}}
if($cleanupFailure){throw ('Owned Access cleanup failed: '+$cleanupFailure)}
foreach($path in $originals.Keys){if((Get-FileHash -LiteralPath $path).Hash -ne $originals[$path]){throw 'A supplied artifact changed'}}
if((Get-FileHash -LiteralPath $baseline).Hash -ne $baselineHash){throw 'Candidate import changed its benchmark source'}
$samples|Export-Csv -LiteralPath (Join-Path $folder 'timings.csv') -NoTypeInformation
$medians=foreach($case in @('Folder stamp','Directory history','Group RSCA')){
 $a=@($samples|Where-Object {$_.Case -eq $case -and $_.Mode -eq 'Baseline'}|Sort-Object Milliseconds)
 $b=@($samples|Where-Object {$_.Case -eq $case -and $_.Mode -eq 'Candidate'}|Sort-Object Milliseconds)
 $index=[int][math]::Floor($Trials/2)
 [pscustomobject]@{Case=$case;BaselineMs=$a[$index].Milliseconds;CandidateMs=$b[$index].Milliseconds;Speedup=$a[$index].Milliseconds/$b[$index].Milliseconds}
}
[pscustomobject]@{Status='PASS';Rows=$Size;Trials=$Trials;ArtifactHashes=$originals;InstalledSource=[bool]$InstallSource;IdenticalReadOutputs=$true;NoReadHistoryWrites=$true;SourceUnchanged=$true;Assertions=$result;Medians=$medians;Scope='Exact input artifact copies; identical synthetic data transferred through production migration; development performance/equivalence only'}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $folder 'verification.json')
$medians|Format-Table -AutoSize
Write-Output ('PASS: batched read equivalence, guards and measurements; evidence: '+$folder)
