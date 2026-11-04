# File-only scratch retention. Never inspect or modify a database through Access.
function GetSwiftScratchEntries([string]$Root){
 $build=Join-Path $Root 'build'
 if(Test-Path -LiteralPath $build){
  Get-ChildItem -LiteralPath $build -Directory -Filter 'source-check-*'
  $tests=Join-Path $build 'tests'
  if(Test-Path -LiteralPath $tests){Get-ChildItem -LiteralPath $tests -Force}
 }
}

function StartSwiftScratchRun([string]$Root){
 $folder=Join-Path $Root 'build/retention/runs'
 [void](New-Item -ItemType Directory -Force -Path $folder)
 $guard=[IO.File]::Open((Join-Path $Root 'build/retention/workflows.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::ReadWrite)
 try {
 $before=@{}
 foreach($item in @(GetSwiftScratchEntries $Root)){$before[$item.FullName]=$item.LastWriteTimeUtc.Ticks}
 $lease=Join-Path $folder ([guid]::NewGuid().ToString()+'.active.json')
 @{PID=$PID;ProcessStartedUtc=(Get-Process -Id $PID).StartTime.ToUniversalTime().ToString('o')}|ConvertTo-Json|Set-Content -LiteralPath $lease
 [pscustomobject]@{Root=$Root;Before=$before;Lease=$lease;Guard=$guard}
 } catch {$guard.Dispose();throw}
}

function CompleteSwiftScratchRun($Run,[string]$Status){
 try {
  $entries=@(foreach($item in @(GetSwiftScratchEntries $Run.Root)){
   if(-not $Run.Before.ContainsKey($item.FullName) -or $Run.Before[$item.FullName] -ne $item.LastWriteTimeUtc.Ticks){
    @{Path=$item.FullName;Ticks=$item.LastWriteTimeUtc.Ticks}
   }
  })
  @{Status=$Status;CompletedUtc=[datetime]::UtcNow.ToString('o');Entries=$entries}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath ($Run.Lease -replace '\.active\.json$','.json')
 } finally {
  Remove-Item -LiteralPath $Run.Lease -ErrorAction SilentlyContinue
  $Run.Guard.Dispose()
 }
}

function InvokeSwiftScratchCleanup {
 [CmdletBinding()]
 param([Parameter(Mandatory=$true)][string]$Root,[switch]$Apply,[string[]]$PreservePath=@(),
       [ValidateRange(1,8760)][int]$SuccessHours=24,[ValidateRange(1,365)][int]$OtherDays=7)
 $Root=[IO.Path]::GetFullPath($Root)
 $build=[IO.Path]::GetFullPath((Join-Path $Root 'build'))
 $tests=Join-Path $build 'tests'
 $records=Join-Path $build 'retention'
 # Fail closed around directory junctions/symlinks, including the allowed roots.
 foreach($path in @($Root,$build,$tests,$records,(Join-Path $records 'runs'))){
  if((Test-Path -LiteralPath $path) -and ((Get-Item -LiteralPath $path -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw ('Retention refuses linked root: '+$path)}
 }
 [void](New-Item -ItemType Directory -Force -Path $records)
 try {$cleanupGuard=[IO.File]::Open((Join-Path $records 'workflows.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)}
 catch {
  Write-Host 'SKIP scratch cleanup: a workflow or cleanup holds the retention lock.'
  return [pscustomobject]@{Status='SKIP';Reason='Retention lock held';DeletedBytes=0}
 }
 try {
 if(Get-Process MSACCESS -ErrorAction SilentlyContinue){
  Write-Host 'SKIP scratch cleanup: Access is running.'
  return [pscustomobject]@{Status='SKIP';Reason='Access running';DeletedBytes=0}
 }
 $states=@{}
 $runs=Join-Path $records 'runs'
 if(Test-Path -LiteralPath $runs){
  foreach($file in @(Get-ChildItem -LiteralPath $runs -Filter '*.active.json')){
   try {
    $lease=Get-Content -LiteralPath $file.FullName -Raw|ConvertFrom-Json
    $owner=Get-Process -Id $lease.PID -ErrorAction SilentlyContinue
    if($owner -and $owner.StartTime.ToUniversalTime().Ticks -eq ([datetime]$lease.ProcessStartedUtc).ToUniversalTime().Ticks){
     Write-Host 'SKIP scratch cleanup: a managed workflow is active.'
     return [pscustomobject]@{Status='SKIP';Reason='Active workflow';DeletedBytes=0}
    }
   } catch {throw ('Cannot verify scratch lease: '+$file.FullName)}
  }
  foreach($file in @(Get-ChildItem -LiteralPath $runs -Filter '*.json'|Where-Object {$_.Name -notlike '*.active.json'}|Sort-Object LastWriteTimeUtc)){
   $record=Get-Content -LiteralPath $file.FullName -Raw|ConvertFrom-Json
   foreach($entry in $record.Entries){
    $prior=$states[[string]$entry.Path]
    $status=$record.Status
    if($prior -and $prior.Ticks -eq [long]$entry.Ticks -and $prior.Status -ne 'PASS'){$status=$prior.Status}
    $states[[string]$entry.Path]=@{Status=$status;Ticks=[long]$entry.Ticks;Completed=([datetime]$record.CompletedUtc).ToUniversalTime()}
   }
  }
 }
 $now=[datetime]::UtcNow;$candidates=[Collections.Generic.List[object]]::new();$skipped=[Collections.Generic.List[object]]::new()
 foreach($item in @(GetSwiftScratchEntries $Root)){
  if($item.Name -like '*.keep'){continue}
  $path=[IO.Path]::GetFullPath($item.FullName)
  $parent=Split-Path $path -Parent
  # An allowlist of generated names leaves manually named/unknown artifacts alone.
  $generated=($item.Name -match '[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}') -or
             (-not $item.PSIsContainer -and $item.Name -match '^[a-z-]+-results\.json$')
  if(-not (($parent -eq $build -and $item.PSIsContainer -and $item.Name -match '^source-check-[0-9a-f-]{36}$') -or ($parent -eq $tests -and $generated))){continue}
  if(Test-Path -LiteralPath ($path+'.keep')){continue}
  $protected=$false
  foreach($keep in $PreservePath){
   if(-not $keep){continue}
   $keep=[IO.Path]::GetFullPath($keep)
   if($keep -eq $path -or $keep.StartsWith($path+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){$protected=$true}
  }
  if($protected){continue}
  $tree=[Collections.Generic.List[object]]::new();$tree.Add($item)
  $blocked=$false
  # Traverse explicitly; never follow a reparse point, even when nested.
  for($i=0;$i -lt $tree.Count;$i++){
   $node=$tree[$i]
   if($node.Attributes -band [IO.FileAttributes]::ReparsePoint){$blocked=$true;break}
   if($node.Name -eq '.keep' -or $node.Extension -in @('.laccdb','.ldb')){$blocked=$true;break}
   if($node.PSIsContainer){foreach($child in @(Get-ChildItem -LiteralPath $node.FullName -Force)){$tree.Add($child)}}
  }
  if($blocked){$skipped.Add(@{Path=$path;Reason='Pinned, linked or lock file'});continue}
  $latest=($tree|Measure-Object LastWriteTimeUtc -Maximum).Maximum
  $status='UNKNOWN';$state=$states[$path]
  if($state -and $state.Ticks -eq $item.LastWriteTimeUtc.Ticks -and $latest -le $state.Completed){$status=$state.Status}
  $ageHours=if($status -eq 'PASS'){$SuccessHours}else{$OtherDays*24}
  $reference=$latest
  if($state -and $state.Completed -gt $reference){$reference=$state.Completed}
  if(($now-$reference).TotalHours -lt $ageHours){continue}
  $files=@($tree|Where-Object {-not $_.PSIsContainer})
  $reason=$status+' scratch older than '+$ageHours+' hours'
  $candidates.Add([pscustomobject]@{Path=$path;Reason=$reason;Files=$files;Tree=$tree;Bytes=($files|Measure-Object Length -Sum).Sum})
 }
 [void](New-Item -ItemType Directory -Force -Path $records)
 $manifest=Join-Path $records ('cleanup-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid()+'.jsonl')
 $writer=[IO.StreamWriter]::new($manifest,$false,[Text.UTF8Encoding]::new($false))
 $totalBytes=($candidates|Measure-Object Bytes -Sum).Sum
 Write-Host ('Running scratch cleanup '+$(if($Apply){'apply'}else{'preview'})+': '+$candidates.Count+' candidates / '+[math]::Round($totalBytes/1GB,2)+' GiB; hashing: '+$manifest)
 $progressAt=[datetime]::UtcNow;$processed=0
 try {
 $deletedBytes=0L;$deletedCount=0;$plannedBytes=0L
 foreach($candidate in $candidates){
  $plannedBytes+=$candidate.Bytes
  $handles=[Collections.Generic.List[object]]::new()
  try {
   if($candidate.Tree[0].PSIsContainer){$writer.WriteLine((@{Path=$candidate.Path;Reason=$candidate.Reason;Type='Directory';Action=$(if($Apply){'Candidate'}else{'Preview'})}|ConvertTo-Json -Compress))}
   # Keep exclusive handles during hashing and deletion (Delete sharing only).
   # Existing readers/writers make the entire candidate ineligible.
   foreach($file in $candidate.Files){
    $stream=[IO.File]::Open($file.FullName,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Delete)
    $handles.Add($stream)
    $sha=[Security.Cryptography.SHA256]::Create()
    try{$hash=([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-','')}finally{$sha.Dispose()}
    $writer.WriteLine((@{Path=$file.FullName;Reason=$candidate.Reason;Bytes=$stream.Length;SHA256=$hash;Action=$(if($Apply){'Candidate'}else{'Preview'})}|ConvertTo-Json -Compress))
   }
   $writer.Flush()
   if($Apply){
    if(Get-Process MSACCESS -ErrorAction SilentlyContinue){throw 'Access started during cleanup'}
    # Revalidate every resolved target immediately before recursive removal.
    $target=[IO.Path]::GetFullPath($candidate.Path)
    if((Split-Path $target -Parent) -notin @($build,$tests)){throw 'Deletion target escapes allowed roots'}
    if(Test-Path -LiteralPath ($target+'.keep')){throw 'Candidate was pinned after inventory'}
    $current=@(Get-Item -LiteralPath $target -Force)+@(if((Get-Item -LiteralPath $target).PSIsContainer){Get-ChildItem -LiteralPath $target -Recurse -Force})
    if($current.Count -ne $candidate.Tree.Count -or @($current|Where-Object {($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -or $_.Name -eq '.keep' -or $_.Extension -in @('.laccdb','.ldb')}).Count){throw 'Candidate changed after inventory'}
    foreach($node in $current){$original=@($candidate.Tree|Where-Object FullName -eq $node.FullName);if($original.Count -ne 1 -or $original[0].LastWriteTimeUtc -ne $node.LastWriteTimeUtc){throw 'Candidate changed after inventory'}}
    Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction Stop
    $deletedBytes+=$candidate.Bytes;$deletedCount++
    $writer.WriteLine((@{Path=$target;Action='Deleted';Bytes=$candidate.Bytes;CompletedUtc=[datetime]::UtcNow.ToString('o')}|ConvertTo-Json -Compress))
   }
  } catch {
   $skipped.Add(@{Path=$candidate.Path;Reason=$_.Exception.Message})
   $writer.WriteLine((@{Path=$candidate.Path;Action='Skipped';Reason=$_.Exception.Message}|ConvertTo-Json -Compress))
  }
  finally {foreach($handle in $handles){$handle.Dispose()}}
  $processed++
  if(([datetime]::UtcNow-$progressAt).TotalSeconds -ge 20){Write-Host ('Running scratch cleanup: '+$processed+'/'+$candidates.Count+' candidates processed');$progressAt=[datetime]::UtcNow}
 }
 } finally {$writer.Dispose()}
 $summary=[pscustomobject]@{Status=$(if($Apply){'APPLIED'}else{'PREVIEW'});Candidates=$candidates.Count;PlannedBytes=$plannedBytes;Deleted=$deletedCount;DeletedBytes=$deletedBytes;Skipped=@($skipped.ToArray());Manifest=$manifest}
 $summary|ConvertTo-Json -Depth 5|Set-Content -LiteralPath ($manifest+'.summary.json')
 Write-Host ('Scratch cleanup '+$summary.Status+': '+$deletedCount+' removed; '+[math]::Round($deletedBytes/1GB,2)+' GiB; '+$skipped.Count+' skipped; manifest: '+$manifest)
 return $summary
 } finally {$cleanupGuard.Dispose()}
}
