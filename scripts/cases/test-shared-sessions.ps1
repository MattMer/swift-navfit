param([Parameter(Mandatory=$true)][string]$BuildPath,[int]$Workers=8)
# Several people on one data file: eight Access sessions, each with its own copy of the
# application, share one data file and work at the same moments (SharedSessionHarness).
# The data file is reached through the \\localhost administrative share when this
# computer allows it, so file locking goes through SMB as on a network drive; otherwise
# a local folder is used and the output says so.
#  1 All start at once: none is told the data file is missing.
#  2 Each saves its own report four times at a working pace, then all eight at the
#    same second: every save lands (a save waits its turn), each with the next number,
#    and the last is the report's saved state. (Each save replaces the working revision
#    before it, so the count of revisions stays the same.)
#  3 Four save one report from the same version: exactly one wins, the rest are refused.
#    Two save one member and two one reporting senior the same way: one wins each.
#  4 Each sees the other seven as connected; one leaves and is gone at once; when all
#    close, no session is left listed.
#  5 A backup made while the others save is verified, or refused with no file left.
#  6 Compacting is refused while others are connected and names them; once all have
#    closed, it compacts with every row kept and a verified backup beside the file.
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$folder=Join-Path $root ('build/tests/shared-sessions-'+[guid]::NewGuid())
[void](New-Item -ItemType Directory -Path $folder)
$drive=$folder.Substring(0,1)
$unc='\\localhost\'+$drive+'$'+$folder.Substring(2)
$dataFolder=if(Test-Path -LiteralPath $unc){$unc}else{$folder}
$route=if($dataFolder -eq $unc){'network path '+$unc}else{'local folder (\\localhost share not reachable)'}
$data=Join-Path $dataFolder 'Swift NAVFIT data.accdb'
$coordinator=Join-Path $folder 'app0.accdb'
Copy-Item -LiteralPath $BuildPath -Destination $coordinator
$checks=0
function Check([bool]$ok,[string]$message){if(-not $ok){throw $message};$script:checks++}

$app=$null
try{
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($coordinator,$true)
    $app.LoadFromText(5,'SharedSessionHarness',(Join-Path $root 'tests/SharedSessionHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $setup=[string]$app.Eval('SharedSetup("'+$data+'")')
    if(-not $setup.StartsWith('PASS:')){throw $setup}
    $ids=$setup.Substring(5)
}finally{if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)};$app=$null}

$engine=New-Object -ComObject DAO.DBEngine.120
function Scalar([string]$sql){$db=$engine.OpenDatabase($data,$false,$true);try{$r=$db.OpenRecordset($sql,4);try{$r.Fields.Item(0).Value}finally{$r.Close()}}finally{$db.Close()}}
$parts=$ids.Split('|');$everyone=[int]$parts[0];$member=[int]$parts[1];$senior=[int]$parts[2]
$ownReports=@(1..$Workers|ForEach-Object {[int]$parts[2+$_].Split(';')[1]})
$before=@{};$beforeTop=@{}
foreach($r in $ownReports+$everyone){$before[$r]=[int](Scalar ('SELECT Count(*) FROM ReportRevision WHERE ReportID='+$r));$beforeTop[$r]=[int](Scalar ('SELECT Max(RevisionNumber) FROM ReportRevision WHERE ReportID='+$r))}
$memberVersion=[int](Scalar ('SELECT VersionCounter FROM CommandMember WHERE MemberID='+$member))
$seniorVersion=[int](Scalar ('SELECT VersionCounter FROM ReportingSeniorRecord WHERE SeniorID='+$senior))

#Every worker gets its own copy of the application, already linked to the data file.
#The shared start time leaves room for eight copies of Access to open.
$start=[DateTime]::Now.AddSeconds(45).ToOADate().ToString([Globalization.CultureInfo]::InvariantCulture)
$procs=@()
foreach($w in 1..$Workers){
    $copy=Join-Path $folder ('app'+$w+'.accdb');Copy-Item -LiteralPath $coordinator -Destination $copy
    $out=Join-Path $folder ('worker'+$w+'.txt')
    $procs+=Start-Process -FilePath 'powershell.exe' -PassThru -WindowStyle Hidden -RedirectStandardError (Join-Path $folder ('worker'+$w+'.err')) -ArgumentList @('-NoProfile','-File',('"'+(Join-Path $PSScriptRoot 'shared-session-worker.ps1')+'"'),'-App',('"'+$copy+'"'),'-Worker',$w,'-Ids',$ids,'-Start',$start,'-Out',('"'+$out+'"'))
}
foreach($p in $procs){if(-not $p.WaitForExit(600000)){throw 'A worker did not finish within 10 minutes'}}

$results=@{}
foreach($w in 1..$Workers){
    $file=Join-Path $folder ('worker'+$w+'.txt')
    if(-not (Test-Path -LiteralPath $file)){throw ('Worker '+$w+' wrote no result: '+(Get-Content -Raw -LiteralPath (Join-Path $folder ('worker'+$w+'.err')) -ErrorAction SilentlyContinue))}
    $results[$w]=[IO.File]::ReadAllText($file)
}
foreach($w in 1..$Workers){
    $text=$results[$w]
    Check ($text -notmatch 'error:') ('Worker '+$w+' hit an error: '+$text)
    Check ($text -match 'ready=True') ('Worker '+$w+' could not use the data file at startup: '+$text)
    Check ($text -match 'hold=held') ('Worker '+$w+' did not open the data file with record-level locking: '+$text)
    Check ($text -match 'own=done' -and $text -notmatch 'own-save') ('Worker '+$w+' had a save refused on its own report: '+$text)
    Check ($text -match ('others='+($Workers-1)+'\b')) ('Worker '+$w+' did not see the other '+($Workers-1)+': '+$text)
}
Write-Output ('Save seconds (four at a working pace, then the same-second save): '+((1..$Workers|ForEach-Object {'w'+$_+' '+[regex]::Match($results[$_],'seconds=([^
]*)').Groups[1].Value}) -join '; '))
Check ($results[1] -match ('after-leave='+($Workers-2)+'\b')) ('A session that left was still listed: '+$results[1])
foreach($w in 1..$Workers){
    $r=$ownReports[$w-1]
    $top=[int](Scalar ('SELECT Max(RevisionNumber) FROM ReportRevision WHERE ReportID='+$r))
    Check ($top -eq $beforeTop[$r]+5) ('Report '+$r+' reached revision '+$top+', expected '+($beforeTop[$r]+5)+' after five saves')
    $count=[int](Scalar ('SELECT Count(*) FROM ReportRevision WHERE ReportID='+$r))
    $distinct=[int](Scalar ('SELECT Count(*) FROM (SELECT DISTINCT RevisionNumber FROM ReportRevision WHERE ReportID='+$r+')'))
    Check ($count -eq $distinct) ('Report '+$r+' has two revisions with the same number')
    Check ([string](Scalar ('SELECT Title FROM Report WHERE ReportID='+$r)) -eq ('Own report '+$w+' burst')) ('Report '+$r+' does not hold its last save')
}
$won=@(1..4|Where-Object {$results[$_] -match 'contest=won'}).Count
$stale=@(1..4|Where-Object {$results[$_] -match 'contest=stale'}).Count
Check ($won -eq 1 -and $stale -eq 3) ('Shared report: '+$won+' saves won and '+$stale+' were refused; expected 1 and 3')
Check ([int](Scalar ('SELECT Max(RevisionNumber) FROM ReportRevision WHERE ReportID='+$everyone)) -eq $beforeTop[$everyone]+1) 'Shared report got more than one new revision'
$winner=@(1..4|Where-Object {$results[$_] -match 'contest=won'})[0]
Check ([string](Scalar ('SELECT DraftNotes FROM Report WHERE ReportID='+$everyone)) -eq ('Saved by worker '+$winner)) 'Shared report does not hold the winning save'
Check (@(5,6|Where-Object {$results[$_] -match 'contest=won'}).Count -eq 1) 'Shared member: not exactly one save won'
Check ([int](Scalar ('SELECT VersionCounter FROM CommandMember WHERE MemberID='+$member)) -eq $memberVersion+1) 'Shared member changed more than once'
Check (@(7,8|Where-Object {$results[$_] -match 'contest=won'}).Count -eq 1) 'Shared reporting senior: not exactly one save won'
Check ([int](Scalar ('SELECT VersionCounter FROM ReportingSeniorRecord WHERE SeniorID='+$senior)) -eq $seniorVersion+1) 'Shared reporting senior changed more than once'
Check ([int](Scalar 'SELECT Count(*) FROM SessionPresence') -eq 0) 'Sessions were still listed after everyone closed'
$backup=Join-Path $folder 'backup-during-saves.accdb'
if($results[2] -match 'backup=done'){Check ((Test-Path -LiteralPath $backup) -and (Test-Path -LiteralPath ($backup+'.verification.json'))) 'The backup made during saves has no file or verification'}
elseif($results[2] -match 'backup=refused'){Check (-not (Test-Path -LiteralPath $backup)) 'A refused backup left its file behind'}
else{Check $false ('The backup during saves failed: '+$results[2])}
Write-Output ('Backup during saves: '+([regex]::Match($results[2],'backup=[^\r\n]*').Value))
Check ($results[1] -match 'compact=refused') ('Compacting was not refused while others were connected: '+$results[1])

#Once everyone has closed, the coordinator compacts the data file.
$rows=[int](Scalar 'SELECT Count(*) FROM ReportRevision')
$app=$null
try{
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($coordinator,$false)
    $compact=[string]$app.Eval('SharedCompact('+[int]$parts[3].Split(';')[0]+')')
}finally{if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)};$app=$null}
Check ($compact -like 'PASS:The data file was compacted from * MB to * MB. A verified backup was made first: *') ('Compact: '+$compact)
Check ([int](Scalar 'SELECT Count(*) FROM ReportRevision') -eq $rows) 'Compacting changed the rows'
Check (@(Get-ChildItem -LiteralPath $folder -Filter 'Swift_NAVFIT_backup_before_compact_*.accdb').Count -eq 1) 'No backup was made before compacting'
Check (@(Get-ChildItem -LiteralPath $folder -Filter '*compact*.accdb' | Where-Object {$_.Name -notlike 'Swift_NAVFIT_backup_before_compact_*'}).Count -eq 0) 'Compacting left a temporary or old copy behind'
Write-Output ('PASS: '+$checks+' shared data file checks with '+$Workers+' sessions at once on a '+$route)
