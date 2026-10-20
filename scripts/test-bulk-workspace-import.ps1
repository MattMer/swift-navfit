param([Parameter(Mandatory=$true)][string]$BuildPath,[string]$BaselinePath,[string]$DemoPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
if(-not $BaselinePath){$BaselinePath=Join-Path $root 'output/Swift_NAVFIT_Blank_2026-10-19-r121b.accdb'}
if(-not $DemoPath){$DemoPath=Join-Path $root 'output/Swift_NAVFIT_Demo_2026-10-19-r121b.accdb'}
$folder=Join-Path $root ('build/tests/bulk-import-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
$source=Join-Path $folder "Sailor's synthetic workspace.accdb";Copy-Item -LiteralPath $DemoPath -Destination $source
$engine=New-Object -ComObject DAO.DBEngine.120;$db=$engine.OpenDatabase($source)
$longNote=('Long note: '+[char]0x03A9+' '+[char]0x6F22+" ' quotation `r`n")*350 -join ''
try{
 $db.Execute("INSERT INTO ReportNote (NoteGuid,ReportID,AuthorProfileID,ContributorName,NoteText,CreatedAt,SourceKind) SELECT TOP 1 '"+[guid]::NewGuid().ToString('B')+"',ReportID,1,'Synthetic benchmark','Placeholder',Now(),'Local' FROM Report",128)
 $rows=$db.OpenRecordset('ReportNote',2);while(-not $rows.EOF){$rows.Edit();$rows.Fields.Item('NoteText').Value=$longNote;$rows.Update();$rows.MoveNext()};$rows.Close()
 # Parent IDs deliberately sort after their children; bulk restoration must not depend on insertion order.
 $db.Execute("INSERT INTO Folder (FolderID,FolderGuid,FolderName) VALUES (8002,'"+[guid]::NewGuid().ToString('B')+"','Synthetic parent')",128)
 $db.Execute("INSERT INTO Folder (FolderID,FolderGuid,FolderName,ParentFolderID) VALUES (8001,'"+[guid]::NewGuid().ToString('B')+"','Synthetic child',8002)",128)
 $db.Execute("INSERT INTO ReviewComment (CommentID,CommentGuid,ReportID,RevisionID,ThreadID,AuthorProfileID,CommentAt,CommentBody) SELECT TOP 1 8001,'"+[guid]::NewGuid().ToString('B')+"',ReportID,RevisionID,Null,AuthorProfileID,CommentAt,CommentBody FROM ReviewComment",128)
 $db.Execute("INSERT INTO ReviewComment (CommentID,CommentGuid,ReportID,RevisionID,ThreadID,AuthorProfileID,CommentAt,CommentBody) SELECT 8002,'"+[guid]::NewGuid().ToString('B')+"',ReportID,RevisionID,8001,AuthorProfileID,CommentAt,CommentBody FROM ReviewComment WHERE CommentID=8001",128)
 # Expand genuine historical revision/value rows without changing the reports' current revisions.
 for($j=1;$j -le 7;$j++){
  $offset=1000*$j;$valueOffset=10000*$j
  $db.Execute('INSERT INTO ReportRevision (RevisionID,ReportID,RevisionNumber,ParentRevisionID,AuthorProfileID,CommittedAt,Reason,SnapshotSchema) SELECT RevisionID+'+$offset+',ReportID,RevisionNumber+'+$offset+',ParentRevisionID,AuthorProfileID,CommittedAt,Reason,SnapshotSchema FROM ReportRevision WHERE RevisionID<1000',128)
  $db.Execute('INSERT INTO RevisionValue (RevisionValueID,RevisionID,FieldKey,ValueText,IsNullValue) SELECT RevisionValueID+'+$valueOffset+',RevisionID+'+$offset+',FieldKey,ValueText,IsNullValue FROM RevisionValue WHERE RevisionValueID<10000',128)
 }
 $rows=$db.OpenRecordset('SELECT TOP 1 NoteText FROM ReportNote',4);if($rows.EOF -or $rows.Fields.Item(0).Value -cne $longNote){throw 'Long note fixture not saved'};$rows.Close()
 $rows=$db.OpenRecordset('SELECT Count(*) FROM RevisionValue',4);$valueCount=[int]$rows.Fields.Item(0).Value;$rows.Close()
}finally{$db.Close()}
$sourceHash=(Get-FileHash -LiteralPath $source).Hash;$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash;$baselineHash=(Get-FileHash -LiteralPath $BaselinePath).Hash
$app=New-Object -ComObject Access.Application;$results=@()
function OpenTarget([string]$template,[string]$label){
 $target=Join-Path $folder ($label+'.accdb');Copy-Item -LiteralPath $template -Destination $target
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
 foreach($name in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
 $app.LoadFromText(5,'WorkspaceImportHarness',(Join-Path $root 'tests/WorkspaceImportHarness.bas'))
 [void]$app.Eval('ConfigureWorkspace("SYNTHETIC BULK BENCHMARK","Tester")')
 return $target
}
try{
 foreach($run in @(@('row',1),@('bulk',1),@('bulk',2),@('row',2))){
  $kind=$run[0];$label=$kind+'-'+$run[1];$template=if($kind -eq 'row'){$BaselinePath}else{$BuildPath};$target=OpenTarget $template $label
  $clock=[Diagnostics.Stopwatch]::StartNew();$result=$app.Eval('TryWorkspaceImport("'+$source+'")');$clock.Stop()
  if($result -notlike 'PASS:*'){throw $result}
  if($app.Eval('DLookup("ParentFolderID","Folder","FolderID=8001")') -ne 8002 -or $app.Eval('DLookup("ThreadID","ReviewComment","CommentID=8002")') -ne 8001){throw 'Deferred self references changed'}
  if(-not $app.Eval('ProbeImportedCounter()')){throw 'Next imported AutoNumber would collide'}
  if($app.Eval('DCount("*","RevisionValue")') -ne $valueCount){throw 'Historical revision values lost'}
  if($app.Eval('DCount("*","MSysObjects","Name Like ''zSwiftMigration*''")') -ne 0){throw 'Temporary migration query remains'}
  $app.CloseCurrentDatabase()
  $db=$engine.OpenDatabase($target,$false,$true);try{$rows=$db.OpenRecordset('SELECT TOP 1 NoteText FROM ReportNote',4);if($rows.Fields.Item(0).Value -cne $longNote){throw 'Long/Unicode note changed or truncated'};$rows.Close()}finally{$db.Close()}
  $results+=[pscustomobject]@{Engine=$kind;Run=$run[1];Seconds=$clock.Elapsed.TotalSeconds;Result=$result;Target=$target;Reopened=$true}
  Write-Output ('PASS '+$label+': '+[Math]::Round($clock.Elapsed.TotalSeconds,3)+' seconds, '+$valueCount+' revision values; long Unicode memo and deferred references exact')
 }
 # Measure the actual Settings/Help Apply path as well, including visible progress updates.
 $target=OpenTarget $BuildPath 'bulk-ui'
 if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$source+'")')){throw 'Bulk UI preview failed'}
 $clock=[Diagnostics.Stopwatch]::StartNew();$applied=$app.Eval('UXApplySwiftWorkspace()');$clock.Stop();$uiSeconds=$clock.Elapsed.TotalSeconds
 if(-not $applied -or $app.Eval('UXWorkspaceImportBusy()')){throw 'Bulk UI apply failed'}
 $uiLog=$app.Eval('Forms!frmUXWorkspaceImport!txtProgress.Value');$uiLog|Set-Content (Join-Path $folder 'ui-progress.log')
 foreach($phase in @('Bulk copying finished','Field verification finished','Import committed successfully')){if($uiLog -notlike ('*'+$phase+'*')){throw ('Missing bulk UI phase '+$phase)}}
 $app.CloseCurrentDatabase()
 Write-Output ('PASS: actual UI import including progress: '+[Math]::Round($uiSeconds,3)+' seconds')
 # Unlike the previous row importer, bulk import also supports a reply whose parent sorts after it.
 $edgeSource=Join-Path $folder 'out-of-order-threads.accdb';Copy-Item -LiteralPath $source -Destination $edgeSource
 $db=$engine.OpenDatabase($edgeSource)
 try{
  $db.Execute("INSERT INTO ReviewComment (CommentID,CommentGuid,ReportID,RevisionID,ThreadID,AuthorProfileID,CommentAt,CommentBody) SELECT TOP 1 9002,'"+[guid]::NewGuid().ToString('B')+"',ReportID,RevisionID,Null,AuthorProfileID,CommentAt,CommentBody FROM ReviewComment",128)
  $db.Execute("INSERT INTO ReviewComment (CommentID,CommentGuid,ReportID,RevisionID,ThreadID,AuthorProfileID,CommentAt,CommentBody) SELECT 9001,'"+[guid]::NewGuid().ToString('B')+"',ReportID,RevisionID,9002,AuthorProfileID,CommentAt,CommentBody FROM ReviewComment WHERE CommentID=9002",128)
 }finally{$db.Close()}
 $edgeHash=(Get-FileHash -LiteralPath $edgeSource).Hash
 $target=OpenTarget $BuildPath 'out-of-order-thread-target'
 $edgeResult=$app.Eval('TryWorkspaceImport("'+$edgeSource+'")')
 if($edgeResult -notlike 'PASS:*' -or $app.Eval('DLookup("ThreadID","ReviewComment","CommentID=9001")') -ne 9002){throw ('Out of order thread failed: '+$edgeResult)}
 $app.CloseCurrentDatabase()
 if((Get-FileHash -LiteralPath $edgeSource).Hash -ne $edgeHash){throw 'Edge source modified'}
 Write-Output 'PASS: out-of-order comment parent restored by bulk update'
 # A deliberately shortened destination memo must fail after earlier tables were copied, rolling back bootstrap and data.
 $target=OpenTarget $BuildPath 'reject-truncated-memo';$db=$app.CurrentDb();$db.Execute('ALTER TABLE ReportNote ALTER COLUMN NoteText TEXT(255)',128);$db.Close()
 $result=$app.Eval('TryWorkspaceImport("'+$source+'")')
 if($result -notlike 'REJECTED:*' -or $result -notmatch 'Copy ReportNote|Verify ReportNote' -or $app.Eval('DCount("*","Report")') -ne 0 -or $app.Eval('DLookup("CommandName","Workspace")') -ne 'SYNTHETIC BULK BENCHMARK'){throw 'Bulk failure did not roll back all copied data/bootstrap'}
 $app.CloseCurrentDatabase()
 Write-Output 'PASS: oversized memo rejected with all target data/setup rolled back'
 # Force a late relationship-update failure, then retry on that same empty target.
 $orphanSource=Join-Path $folder 'orphan-comment-source.accdb';Copy-Item -LiteralPath $edgeSource -Destination $orphanSource
 $db=$engine.OpenDatabase($orphanSource)
 try{$db.Relations.Delete('FK_CommentThread');$db.Execute('UPDATE ReviewComment SET ThreadID=999999 WHERE CommentID=9001',128)}finally{$db.Close()}
 $orphanHash=(Get-FileHash -LiteralPath $orphanSource).Hash
 $target=OpenTarget $BuildPath 'reject-orphan-comment-target'
 $result=$app.Eval('TryWorkspaceImport("'+$orphanSource+'")')
 if($result -notlike 'REJECTED:*Restore ReviewComment.ThreadID*' -or $app.Eval('DCount("*","Report")') -ne 0 -or $app.Eval('DLookup("CommandName","Workspace")') -ne 'SYNTHETIC BULK BENCHMARK' -or $app.Eval('DCount("*","MSysObjects","Name Like ''zSwiftMigration*''")') -ne 0){throw ('Late rollback/query cleanup failed: '+$result)}
 $retry=$app.Eval('TryWorkspaceImport("'+$source+'")');if($retry -notlike 'PASS:*'){throw ('Retry after rollback failed: '+$retry)}
 $app.CloseCurrentDatabase()
 if((Get-FileHash -LiteralPath $orphanSource).Hash -ne $orphanHash){throw 'Invalid source was modified'}
 Write-Output 'PASS: late relationship failure rolled back setup/data, removed queries, and allowed a valid retry'
 if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash -or (Get-FileHash -LiteralPath $BuildPath).Hash -ne $buildHash -or (Get-FileHash -LiteralPath $BaselinePath).Hash -ne $baselineHash){throw 'Benchmark source or release modified'}
 $old=($results|Where-Object {$_.Engine -eq 'row'}|Measure-Object Seconds -Average).Average;$new=($results|Where-Object {$_.Engine -eq 'bulk'}|Measure-Object Seconds -Average).Average
 $evidence=@{Status='PASS';Source=$source;SourceSHA256=$sourceHash;BuildSHA256=$buildHash;BaselineSHA256=$baselineHash;RevisionValues=$valueCount;UISeconds=$uiSeconds;Runs=$results;RowAverageSeconds=$old;BulkAverageSeconds=$new;Speedup=$old/$new;LongUnicodeMemoVerified=$true;OutOfOrderSelfReferencesVerified=$true;TruncationFailureRollbackVerified=$true;LateRelationshipRollbackAndRetryVerified=$true;TemporaryQueriesRemoved=$true;NextAutoNumberVerified=$true;SourcesUnchanged=$true}
 $evidence|ConvertTo-Json -Depth 7|Set-Content (Join-Path $folder 'verification.json')
 Write-Output ('PASS: bulk benchmark and rollback; '+$folder)
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)}
