param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$GroupListBenchmark,[string]$BaselinePath,[switch]$InstallSource,[int]$Groups=40,[int]$Members=10,[int]$Trials=5,[string]$DataSourcePath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
if($GroupListBenchmark){
 if(-not $BaselinePath){$BaselinePath=$BuildPath}
 if($Groups -lt 3 -or $Members -lt 2 -or $Groups*$Members -gt 1000 -or $Trials -lt 3 -or $Trials%2 -eq 0){throw 'Use 3+ groups, 2+ members, at most 1000 reports and an odd 3+ trials'}
 $folder=Join-Path $root ('build/tests/summary-list-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
 $hashes=@{};foreach($path in @($BuildPath,$BaselinePath)){$hashes[[IO.Path]::GetFullPath($path)]=(Get-FileHash -LiteralPath $path).Hash}
 if($DataSourcePath){$DataSourcePath=[IO.Path]::GetFullPath($DataSourcePath);$hashes[$DataSourcePath]=(Get-FileHash -LiteralPath $DataSourcePath).Hash}
 $copies=@{Baseline=(Join-Path $folder 'baseline.accdb');Candidate=(Join-Path $folder 'candidate.accdb')}
 Copy-Item -LiteralPath $BaselinePath -Destination $copies.Baseline;Copy-Item -LiteralPath $BuildPath -Destination $copies.Candidate
 $policy=[IO.File]::ReadAllText((Join-Path $root 'src/modGroupPolicy.bas'));$anchor='    GroupWorkflow = GroupPolicyValue(groupID, "Workflow")'
 if(([regex]::Matches($policy,[regex]::Escape($anchor))).Count -ne 1){throw 'Expected workflow read counter boundary missing or duplicated'}
 $instrumented=Join-Path $folder 'read-counter-modGroupPolicy.bas';[IO.File]::WriteAllText($instrumented,$policy.Replace($anchor,'    SummaryListCountWorkflowLookup'+"`r`n"+$anchor))
 $app=$null;$primary=$null;$cleanup=@();$expected=@{};$counts=@{};$samples=[Collections.Generic.List[object]]::new()
 try{
  $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1
  foreach($mode in @('Baseline','Candidate')){
   Write-Output ('Controller: summary-list benchmark; database: '+$copies[$mode])
   $app.OpenCurrentDatabase($copies[$mode],$true)
   $app.LoadFromText(5,'UXGroupLifecycleHarness',(Join-Path $root 'tests/UXGroupLifecycleHarness.bas'));$app.LoadFromText(5,'modGroupPolicy',$instrumented)
   if($mode -eq 'Candidate' -and $InstallSource){$app.LoadFromText(5,'modUXGroupLifecycle',(Join-Path $root 'src/modUXGroupLifecycle.bas'))}
   $app.DoCmd.RunCommand(126)
   if($DataSourcePath){
    $result=$app.Eval('MigrateWorkspace("'+$DataSourcePath+'",False,False,"'+$hashes[$DataSourcePath]+'")');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}
    [void]$app.Eval('SwitchContributor(1)')
    if($app.Eval('DLookup("CommandName","Workspace")') -cne 'SYNTHETIC SUMMARY LIST' -or $app.Eval('DCount("*","Report")') -ne $Groups*$Members -or $app.Eval('DCount("*","SummaryGroup")') -ne $Groups+8){throw 'Data source is not the matching synthetic summary-list fixture'}
   }else{$result=$app.Eval('SeedSummaryListBenchmark('+$Groups+','+$Members+')');Write-Output $result;if($result -notlike 'PASS:*'){throw $result}}
   $counts[$mode]=[string]$app.Eval('SummaryListDataCounts()');$app.CloseCurrentDatabase()
  }
  for($trial=1;$trial -le $Trials;$trial++){
   $order=if($trial%2){@('Baseline','Candidate')}else{@('Candidate','Baseline')}
   foreach($mode in $order){
    $app.OpenCurrentDatabase($copies[$mode],$true);[void]$app.Eval('SwitchContributor(1)')
    foreach($filter in @('All','Active','Archived')){
     $timer=[Diagnostics.Stopwatch]::StartNew();$value=[string]$app.Eval('SummaryListRead("'+$filter+'")');$timer.Stop()
     $lookups=[int]$app.Eval('SummaryListLookupCount()')
     if(-not $expected.ContainsKey($filter)){$expected[$filter]=$value;[IO.File]::WriteAllText((Join-Path $folder ($filter+'-expected.txt')),$value)}elseif($value -cne $expected[$filter]){throw ('Summary list output/order differs: '+$mode+' '+$filter)}
     if($mode -eq 'Candidate' -and $lookups -ne 0){throw 'Candidate summary query still performs scalar workflow lookups'}
     if($mode -eq 'Baseline' -and $lookups -lt 1){throw 'Baseline workflow reads were not observed; no optimization comparison proved'}
     $samples.Add([pscustomobject]@{Mode=$mode;Filter=$filter;Trial=$trial;Milliseconds=$timer.Elapsed.TotalMilliseconds;WorkflowLookups=$lookups;Groups=$Groups;Reports=$Groups*$Members})
    }
    if([string]$app.Eval('SummaryListDataCounts()') -cne $counts[$mode]){throw 'Summary list reads wrote data/history'}
    $app.CloseCurrentDatabase()
   }
   Write-Output ('PASS: summary-list output/order, filters and no writes; alternating trial '+$trial)
  }
 }catch{$primary=$_}finally{
  if($app){try{$app.Quit(2)}catch{$cleanup+=$_.Exception.Message};try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+=$_.Exception.Message}}
  foreach($path in $hashes.Keys){if((Get-FileHash -LiteralPath $path).Hash -ne $hashes[$path]){$cleanup+='Supplied artifact changed: '+$path}}
 }
 if($primary){if($cleanup){Write-Warning ($cleanup -join '; ')};throw $primary};if($cleanup){throw ($cleanup -join '; ')}
 $samples|Export-Csv -LiteralPath (Join-Path $folder 'timings.csv') -NoTypeInformation
 $medians=foreach($filter in @('All','Active','Archived')){$a=@($samples|Where-Object {$_.Filter -eq $filter -and $_.Mode -eq 'Baseline'}|Sort-Object Milliseconds);$b=@($samples|Where-Object {$_.Filter -eq $filter -and $_.Mode -eq 'Candidate'}|Sort-Object Milliseconds);[pscustomobject]@{Filter=$filter;BaselineMs=$a[[int][Math]::Floor($Trials/2)].Milliseconds;CandidateMs=$b[[int][Math]::Floor($Trials/2)].Milliseconds;BaselineWorkflowLookups=$a[0].WorkflowLookups;CandidateWorkflowLookups=$b[0].WorkflowLookups}}
 @{Status='PASS';SourceHashes=$hashes;Groups=$Groups;Reports=$Groups*$Members;Trials=$Trials;InstalledSource=[bool]$InstallSource;DataSourcePath=$DataSourcePath;Medians=$medians;Scope='Identical synthetic displayed fields/IDs; optional fixture reuse uses production migration and hash protection; five alternating native samples by default; workflow counter exists only in test copies; no physical/performance guarantee'}|ConvertTo-Json -Depth 6|Set-Content (Join-Path $folder 'verification.json')
 $medians|Format-Table -AutoSize;Write-Output ('PASS: summary-list benchmark receipt: '+$folder);exit
}
$scratch=Join-Path $root ('build/tests/ux-group-lifecycle-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'UXGroupLifecycleHarness',(Join-Path $root 'tests/UXGroupLifecycleHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestUXGroupLifecycle()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    Write-Output $result
}finally{if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
