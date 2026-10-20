param([Parameter(Mandatory=$true)][string]$BlankPath,[Parameter(Mandatory=$true)][string]$DemoPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$folder=Join-Path $root ('build/tests/message-import-'+[guid]::NewGuid());[void](New-Item -ItemType Directory $folder)
$source=Join-Path $folder 'source.accdb';$target=Join-Path $folder 'target.accdb';Copy-Item $DemoPath $source;Copy-Item $BlankPath $target
$e=New-Object -ComObject DAO.DBEngine.120;$d=$e.OpenDatabase($source)
try {
 $r=$d.OpenRecordset('SELECT TOP 1 ReportID FROM Report WHERE IsArchived=False');$id=$r.Fields.Item(0).Value;$r.Close()
 $r=$d.OpenRecordset('SELECT ActiveProfileID FROM Workspace');$author=$r.Fields.Item(0).Value;$r.Close()
 $parentGuid=[guid]::NewGuid().ToString('B');$childGuid=[guid]::NewGuid().ToString('B');$draftGuid=[guid]::NewGuid().ToString('B')
 $d.Execute("INSERT INTO ReportMessage (MessageID,MessageGuid,ReportID,AuthorProfileID,ContributorName,CreatedAt,MessageBody) VALUES (900,'$parentGuid',$id,$author,'Synthetic author',Now(),'Out of order root')",128)
 $d.Execute("INSERT INTO ReportMessage (MessageID,MessageGuid,ReportID,ParentMessageID,AuthorProfileID,ContributorName,CreatedAt,MessageBody) VALUES (800,'$childGuid',$id,900,$author,'Synthetic author',Now(),'Out of order reply')",128)
 $d.Execute("INSERT INTO ReportNoteDraft (DraftGuid,ReportID,AuthorProfileID,NoteText,CheckpointAt,VersionCounter,ParentMessageID) VALUES ('$draftGuid',$id,$author,'Private reply checkpoint',Now(),1,900)",128)
} finally {$d.Close()}
$hash=(Get-FileHash $source).Hash;$app=New-Object -ComObject Access.Application
try {
 $app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
 $result=$app.Eval('MigrateWorkspace("'+$source+'")');if($result -notlike 'PASS:*'){throw $result}
 if($app.Eval('DLookup("ParentMessageID","ReportMessage","MessageID=800")') -ne 900){throw 'Reply relationship lost'}
 $expr='DLookup("ParentMessageID","ReportNoteDraft","DraftGuid=' + [char]39 + $draftGuid + [char]39 + '")';if($app.Eval($expr) -ne 900){throw 'Draft relationship lost'}
 if($app.Eval('MigrateReportMessages()') -ne 0){throw 'Idempotent conversion duplicated messages'}
 $next=$app.Eval('AddReportMessage('+$id+',"Next message identity")');if($next -le 900){throw 'Message AutoNumber not advanced'}
 $app.CloseCurrentDatabase()
} finally {$app.Quit(2)}
if((Get-FileHash $source).Hash -ne $hash){throw 'Source changed'}
Write-Output 'PASS: schema26 conversations, reversed parent IDs, private reply drafts, exact source fields, idempotency and next AutoNumber'
