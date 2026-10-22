param([Parameter(Mandatory=$true)][string]$BuildPath,[Parameter(Mandatory=$true)][string]$FixtureMapPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$folder=Join-Path $root ('build/tests/audit-message-transfer-'+[guid]::NewGuid())
[void](New-Item -ItemType Directory $folder)
$fixtures=Get-Content -LiteralPath $FixtureMapPath -Raw|ConvertFrom-Json
$engine=New-Object -ComObject DAO.DBEngine.120
$definitions=(Get-Content (Join-Path $root 'src/field-definitions.json') -Raw -Encoding UTF8|ConvertFrom-Json).fields
$results=@();$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
function HasField($table,[string]$name){return @($table.Fields|ForEach-Object {$_.Name}) -contains $name}
# Reflection preserves DAO's Variant setter across mixed string/date/number fields.
# PowerShell's cached COM setter otherwise binds to the first value's CLR type.
function SetField($record,[string]$key,$value){$field=$record.Fields.Item($key);[void]$field.GetType().InvokeMember('Value',[Reflection.BindingFlags]::SetProperty,$null,$field,@($value))}
function AddFixture($db,[int]$schema){
 $tables=@($db.TableDefs|ForEach-Object {$_.Name})
 $r=$db.OpenRecordset('UserProfile',2)
 if($r.EOF){$r.AddNew();SetField $r 'ProfileGuid' ([guid]::NewGuid().ToString('B'));SetField $r 'DisplayName' ('Synthetic transfer actor');$r.Update();$r.Bookmark=$r.LastModified}
 $author=[int]$r.Fields('UserProfileID').Value;$r.Close()
 $r=$db.OpenRecordset('Workspace',2)
 if($r.EOF){$r.AddNew();SetField $r 'WorkspaceID' (1);SetField $r 'WorkspaceGuid' ([guid]::NewGuid().ToString('B'));SetField $r 'CommandName' ('SYNTHETIC TRANSFER');SetField $r 'ActiveProfileID' ($author);$r.Update()}
 $r.Close()
 $r=$db.OpenRecordset('Report',2);$r.AddNew()
 foreach($key in @('ReportGuid','CreatedAt','ModifiedAt','FormKind','Title','MemberName','PeriodStart','PeriodEnd','DraftNotes','WorkflowState','IsArchived')){
  $value=switch($key){'ReportGuid'{[guid]::NewGuid().ToString('B')}'CreatedAt'{Get-Date}'ModifiedAt'{Get-Date}'FormKind'{'EVAL'}'Title'{'Synthetic audit transfer'}'MemberName'{'SYNTHETIC'}'PeriodStart'{[datetime]'2026-01-01'}'PeriodEnd'{[datetime]'2026-12-31'}'DraftNotes'{'Original legacy draft notes'}'WorkflowState'{'Draft'}'IsArchived'{$false}}
  SetField $r $key ($value)
 }
 $r.Update();$r.Bookmark=$r.LastModified;$report=[int]$r.Fields('ReportID').Value;$r.Close()
 $r=$db.OpenRecordset('ReportRevision',2);$r.AddNew();SetField $r 'ReportID' ($report);SetField $r 'RevisionNumber' (1);SetField $r 'AuthorProfileID' ($author);SetField $r 'CommittedAt' (Get-Date);SetField $r 'Reason' ('Synthetic source history')
 if(HasField $db.TableDefs('ReportRevision') 'SnapshotSchema'){SetField $r 'SnapshotSchema' ($schema)}
 if(HasField $db.TableDefs('ReportRevision') 'FieldDefinitionVersion'){SetField $r 'FieldDefinitionVersion' (1)}
 $r.Update();$r.Bookmark=$r.LastModified;$revision=[int]$r.Fields('RevisionID').Value;$r.Close()
 $db.Execute(('UPDATE Report SET CurrentRevisionID='+$revision+' WHERE ReportID='+$report),128)
 $source=$db.OpenRecordset('SELECT * FROM Report WHERE ReportID='+$report,4);$r=$db.OpenRecordset('RevisionValue',2)
 foreach($field in $definitions){if(HasField $db.TableDefs('Report') $field.key){
  $value=$source.Fields($field.key).Value;$nullValue=$null -eq $value -or [Convert]::IsDBNull($value)
  $r.AddNew();SetField $r 'RevisionID' ($revision);SetField $r 'FieldKey' ($field.key);SetField $r 'IsNullValue' ($nullValue)
  if(-not $nullValue){if($value -is [datetime]){$value=$value.ToString('yyyy-MM-dd')};SetField $r 'ValueText' ([string]$value)};$r.Update()
 }}
 $r.Close();$source.Close()
 if($tables -contains 'ReportHead'){$db.Execute(('INSERT INTO ReportHead (ReportID,RevisionID) VALUES ('+$report+','+$revision+')'),128)}
 $r=$db.OpenRecordset('SELECT IIf(Max(CommentID) Is Null,0,Max(CommentID)) FROM ReviewComment',4);$parent=[int]$r.Fields.Item(0).Value+1000;$r.Close()
 foreach($entry in @(@($parent,'Later-numbered legacy root'),@(($parent-1),'Out of order legacy reply'))){
  $r=$db.OpenRecordset('ReviewComment',2);$r.AddNew();SetField $r 'CommentID' ($entry[0]);SetField $r 'CommentGuid' ([guid]::NewGuid().ToString('B'));SetField $r 'ReportID' ($report);SetField $r 'RevisionID' ($revision);SetField $r 'AuthorProfileID' ($author);SetField $r 'CommentAt' ([datetime]'2024-02-29');SetField $r 'CommentBody' ($entry[1]);if($entry[0] -ne $parent){SetField $r 'ThreadID' ($parent)};$r.Update();$r.Close()
 }
 $r=$db.OpenRecordset('ReviewCommentEvent',2);$r.AddNew();SetField $r 'ThreadID' ($parent);SetField $r 'AuthorProfileID' ($author);SetField $r 'EventAt' ([datetime]'2024-03-01');SetField $r 'EventKind' ('Resolved');$r.Update();$r.Close()
 if($tables -contains 'ReportNote'){$r=$db.OpenRecordset('ReportNote',2);$r.AddNew();SetField $r 'NoteGuid' ([guid]::NewGuid().ToString('B'));SetField $r 'ReportID' ($report);SetField $r 'ContributorName' (('N'*160));SetField $r 'CreatedAt' ([datetime]'2024-02-29');SetField $r 'NoteText' ('Preserved historical note');SetField $r 'SourceKind' ('Legacy');$r.Update();$r.Close()}
 $draftGuid=[guid]::NewGuid().ToString('B');$message=0
 if($tables -contains 'ReportMessage'){
  $r=$db.OpenRecordset('SELECT IIf(Max(MessageID) Is Null,0,Max(MessageID)) FROM ReportMessage',4);$message=[int]$r.Fields.Item(0).Value+1000;$r.Close()
  foreach($entry in @(@($message,'Canonical source root'),@(($message-1),'Canonical source reply'))){
   $r=$db.OpenRecordset('ReportMessage',2);$r.AddNew();SetField $r 'MessageID' ($entry[0]);SetField $r 'MessageGuid' ([guid]::NewGuid().ToString('B'));SetField $r 'ReportID' ($report);SetField $r 'AuthorProfileID' ($author);SetField $r 'ContributorName' ('Synthetic author');SetField $r 'CreatedAt' ([datetime]'2024-02-29');SetField $r 'MessageBody' ($entry[1]);if($entry[0] -ne $message){SetField $r 'ParentMessageID' ($message)};$r.Update();$r.Close()
  }
 }
 if($tables -contains 'ReportNoteDraft'){$r=$db.OpenRecordset('ReportNoteDraft',2);$r.AddNew();SetField $r 'DraftGuid' ($draftGuid);SetField $r 'ReportID' ($report);SetField $r 'AuthorProfileID' ($author);SetField $r 'NoteText' ('Private raw checkpoint');SetField $r 'VersionCounter' (1);SetField $r 'CheckpointAt' ([datetime]'2024-02-29');if($message -gt 0){SetField $r 'ParentMessageID' ($message)};$r.Update();$r.Close()}
 return @{Report=$report;Revision=$revision;Parent=$parent;DraftGuid=$draftGuid;Message=$message;Private=($tables -contains 'ReportNoteDraft');Note=($tables -contains 'ReportNote')}
}
foreach($fixture in $fixtures.PSObject.Properties|Sort-Object {[int]$_.Name}){
 $schema=[int]$fixture.Name;$original=$fixture.Value;$originalHash=(Get-FileHash -LiteralPath $original).Hash
 $source=Join-Path $folder ('schema-'+$schema+'-source.accdb');$target=Join-Path $folder ('schema-'+$schema+'-target.accdb')
 Copy-Item -LiteralPath $original -Destination $source;Copy-Item -LiteralPath $BuildPath -Destination $target
 $db=$engine.OpenDatabase($source)
 try{$r=$db.OpenRecordset('SELECT SchemaVersion FROM AppMetadata',4);if([int]$r.Fields.Item(0).Value -ne $schema){throw 'Fixture must retain its real schema'};$r.Close();$ids=AddFixture $db $schema}finally{$db.Close()}
 $sourceHash=(Get-FileHash -LiteralPath $source).Hash;$app=New-Object -ComObject Access.Application
 try {
  $app.AutomationSecurity=1;$app.OpenCurrentDatabase($target,$true)
  [void]$app.Eval('ConfigureWorkspace("SYNTHETIC BOOTSTRAP","Tester")')
  if(-not $app.Eval('UXPreviewSwiftWorkspace("'+$source+'")')){throw 'Settings preview failed'}
  if($schema -eq 21){
   $targetDB=$app.CurrentDb();$targetDB.TableDefs('AuditEvent').Fields('EventKind').ValidationRule="='Reject audit migration'";$targetDB.Close()
   if($app.Eval('UXApplySwiftWorkspace()')){throw 'Injected late audit failure was accepted'}
   if($app.Eval('DCount("*","Report")') -ne 0 -or $app.Eval('DCount("*","AuditEvent")') -ne 0 -or $app.Eval('DCount("*","AuditActor")') -ne 0 -or $app.Eval('DLookup("CommandName","Workspace")') -ne 'SYNTHETIC BOOTSTRAP'){throw 'Late audit migration failed to roll back'}
   $targetDB=$app.CurrentDb();$targetDB.TableDefs('AuditEvent').Fields('EventKind').ValidationRule='';$targetDB.Close()
  }
  if(-not $app.Eval('UXApplySwiftWorkspace()')){throw $app.Eval('Forms!frmUXWorkspaceImport!lblStatus.Caption')}
  if($app.Eval('MigrateReportMessages()') -ne 0){throw 'Conversion was not idempotent'}
  $rootID=$app.Eval('DLookup("MessageID","ReportMessage","LegacyCommentID='+$ids.Parent+'")')
  $childID=$app.Eval('DLookup("MessageID","ReportMessage","LegacyCommentID='+($ids.Parent-1)+'")')
  if($app.Eval('ReportMessageRoot('+$childID+','+$ids.Report+')') -ne $rootID){throw 'Legacy reply relationship lost'}
  if($ids.Note -and $app.Eval('DCount("*","ReportMessage","Len(ContributorName)=160")') -eq 0){throw 'Legacy contributor name truncated'}
  if($ids.Private){$expr='DLookup("NoteText","ReportNoteDraft","DraftGuid='''+$ids.DraftGuid+'''")';if($app.Eval($expr) -ne 'Private raw checkpoint'){throw 'Private draft lost'}}
  if($ids.Message -gt 0){
   if($app.Eval('DLookup("ParentMessageID","ReportMessage","MessageID='+($ids.Message-1)+'")') -ne $ids.Message){throw 'Canonical reversed parent lost'}
   $expr='DLookup("ParentMessageID","ReportNoteDraft","DraftGuid='''+$ids.DraftGuid+'''")';if($app.Eval($expr) -ne $ids.Message){throw 'Private reply context lost'}
  }
  if($schema -eq 29 -and $app.Eval('DCount("*","ReportNote","LegacyReportID='+$ids.Report+'")') -ne 0){throw 'Modern DraftNotes was implicitly published'}
  $audit=$app.Eval('DCount("*","AuditEvent")');[void]$app.Run('MigrateAuditHistory');if($app.Eval('DCount("*","AuditEvent")') -ne $audit){throw 'Audit backfill duplicated events'}
  $app.CloseCurrentDatabase();$app.OpenCurrentDatabase($target,$true)
  if($app.Eval('DCount("*","AuditEvent")') -ne $audit){throw 'Audit identities lost on reopen'}
 } finally {$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 if((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash -or (Get-FileHash -LiteralPath $original).Hash -ne $originalHash){throw 'Source/template changed'}
 $results+=@{Schema=$schema;Source=$source;SourceSHA256=$sourceHash;Original=$original;OriginalSHA256=$originalHash;Target=$target;Status='PASS';AuditRows=$audit;LateAuditRollbackTested=($schema -eq 21)}
 Write-Output ('PASS: actual schema '+$schema+' Settings import, original rows/fields, legacy/current replies, private drafts, actor/event identities, idempotency and reopen')
}
if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $buildHash){throw 'Build changed'}
$results|ConvertTo-Json -Depth 8|Set-Content (Join-Path $folder 'verification.json') -Encoding UTF8
Write-Output ('PASS: '+$results.Count+' real schema fixtures; evidence: '+$folder)
