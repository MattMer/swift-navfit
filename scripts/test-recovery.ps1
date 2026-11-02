param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource,[string]$CompiledPath)
# Backup and recovery suite: each case owns its disposable copy and sessions.
# -InstallSource reloads modBackupUI only in the backup case.
# -CompiledPath adds the ACCDE backup/independent comparison in that case.
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash

Invoke-SuiteCase 'backup-restore' {

$folder=Join-Path $root ('build/tests/backup-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
$source=Join-Path $folder "O'Brien & `$literal [source].accdb"
$backup=Join-Path $folder 'copy & $literal [verified].accdb'
$restored=Join-Path $folder 'restored.accdb'
$originals=@{};foreach($path in @($BuildPath,$CompiledPath|Where-Object {$_})){$originals[[IO.Path]::GetFullPath($path)]=(Get-FileHash -LiteralPath $path).Hash}
Copy-Item -LiteralPath $BuildPath -Destination $source
$app=$null;$primary=$null;$cleanup=@();$results=[Collections.Generic.List[string]]::new()
$engine=New-Object -ComObject DAO.DBEngine.120
function AssertBackup([bool]$ok,[string]$message){if(-not $ok){throw $message}}
function Quote([string]$value){return '"'+$value.Replace('"','""')+'"'}
function Rejected([string]$destination,[string]$expected){
 $outcome=$app.Eval('TryWorkspaceBackup('+(Quote $destination)+')')
 AssertBackup ($outcome -like 'REJECTED:*' -and $outcome -match $expected) ('Expected rejection "'+$expected+'", got: '+$outcome)
 $results.Add('Rejected: '+$expected)
}
#Every backup row must equal the workspace row with the same key, field by field.
#The workspace may only gain rows in tables written after the backup (journal, audit).
function CompareFiles([string]$copy,[string]$workspace,[string[]]$mayGrow=@()){
 $a=$engine.OpenDatabase($copy,$false,$true);$b=$engine.OpenDatabase($workspace,$false,$true);$rows=0
 try{
  $names=@(foreach($t in $a.TableDefs){if($t.Name -notlike 'MSys*' -and $t.Name -notlike '~*'){[string]$t.Name}})
  $other=@(foreach($t in $b.TableDefs){if($t.Name -notlike 'MSys*' -and $t.Name -notlike '~*'){[string]$t.Name}})
  AssertBackup (-not (Compare-Object $names $other)) 'Backup table set differs'
  foreach($name in $names){
   $key=@(foreach($i in $a.TableDefs.Item($name).Indexes){if($i.Primary){foreach($f in $i.Fields){[string]$f.Name}}})
   AssertBackup ($key.Count -gt 0) ('Backup table has no primary key: '+$name)
   $live=@{};$s=$b.OpenRecordset('SELECT * FROM ['+$name+']',4)
   try{while(-not $s.EOF){$v=@{};foreach($f in $s.Fields){$v[$f.Name]=$f.Value};$live[(@($key|ForEach-Object {[string]$s.Fields.Item($_).Value})-join '|')]=$v;$s.MoveNext()}}finally{$s.Close()}
   $r=$a.OpenRecordset('SELECT * FROM ['+$name+']',4);$count=0
   try{
    while(-not $r.EOF){
     $id=(@($key|ForEach-Object {[string]$r.Fields.Item($_).Value})-join '|');$match=$live[$id]
     AssertBackup ($null -ne $match) ('Backup row is not in the workspace: '+$name+' '+$id)
     foreach($f in $r.Fields){
      $x=$f.Value;$y=$match[$f.Name]
      $same=if([Convert]::IsDBNull($x) -or [Convert]::IsDBNull($y)){[Convert]::IsDBNull($x) -and [Convert]::IsDBNull($y)}elseif($x -is [string]){[string]::Equals($x,[string]$y,[StringComparison]::Ordinal)}else{[object]::Equals($x,$y)}
      AssertBackup $same ('Backup value differs: '+$name+'.'+$f.Name+' row '+$id)
     }
     $rows++;$count++;$r.MoveNext()
    }
   }finally{$r.Close()}
   AssertBackup ($count -eq $live.Count -or $mayGrow -contains $name) ('Workspace has rows the backup lacks in '+$name)
  }
 }finally{$a.Close();$b.Close()}
 return $rows
}
try {
 Write-Output ('Controller: backup/recovery runner; disposable database: '+$source)
 $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($source,$true)
 if($InstallSource){$app.LoadFromText(5,'modBackupUI',(Join-Path $root 'src/modBackupUI.bas'))}
 $app.LoadFromText(5,'BackupHarness',(Join-Path $root 'tests/BackupHarness.bas'));$app.DoCmd.RunCommand(126)
 #Workspace size against the 2 GB Access limit (L03): below and at the 1.5 GB warning.
 AssertBackup ($app.Eval('WorkspaceSizeText(1048576)') -like 'Workspace file: 1.0 MB of the 2 GB Access limit (0%).') 'Workspace size text below the warning differs'
 AssertBackup (($app.Eval('WorkspaceSizeText(1610612736)') -like 'Warning: * 1.50 GB of the 2 GB Access limit.*') -and -not $app.Eval('WorkspaceSizeNearLimit(1610612735)')) 'Workspace size warning threshold differs'
 $app.DoCmd.OpenForm('frmUXSettings');[void]$app.Eval('UXSettingsActivate()')
 AssertBackup ($app.Eval('Forms!frmUXSettings!uxWorkspaceSize.Caption') -like 'Workspace file: * MB of the 2 GB Access limit (*%).') 'Settings & Help does not show the workspace size'
 AssertBackup ($app.Eval('WorkspaceFileBytes()') -eq (Get-Item -LiteralPath $source).Length) 'Workspace size differs from the file on disk'
 $app.DoCmd.Close(2,'frmUXSettings',2)
 $result=$app.Eval('PrepareBackupCheckpoint()');AssertBackup ($result -like 'PASS:*') $result
 AssertBackup ([bool]$app.Eval('CurrentProject.AllForms("frmEditor").IsLoaded')) 'Editor was not open with pending values'
 $evidence=$app.Eval('TryWorkspaceBackup('+(Quote $backup)+')')
 AssertBackup ($evidence -like '{"Status":"PASS"*') ('Backup failed: '+$evidence)
 AssertBackup (-not $app.Eval('CurrentProject.AllForms("frmEditor").IsLoaded')) 'Backup did not checkpoint and close the editor'
 AssertBackup ($app.Eval('CurrentProject.FullName') -eq $source) 'Access did not stay open on the workspace'
 AssertBackup ($app.Eval('DCount("*","OperationJournal","OperationKind=''Data backup'' AND OperationState=''Completed''")') -eq 1) 'Backup was not journaled'
 $results.Add('Backup while open: editor checkpointed and closed, Access stays open, literal paths, journaled')
 Rejected $backup 'Destination already exists'
 Rejected (Join-Path $folder 'missing/out.accdb') 'Destination folder does not exist'
 Rejected (Join-Path $folder 'wrong-extension.accde') 'new .accdb data file'
 $protected=Join-Path $folder 'protected.accdb';$evidencePath=$protected+'.verification.json'
 [IO.File]::WriteAllText($evidencePath,'SYNTHETIC PREEXISTING VERIFICATION');$evidenceHash=(Get-FileHash -LiteralPath $evidencePath).Hash
 Rejected $protected 'verification file already exists'
 AssertBackup (-not (Test-Path -LiteralPath $protected) -and (Get-FileHash -LiteralPath $evidencePath).Hash -eq $evidenceHash) 'Rejected backup touched preexisting verification evidence'
 $linkedTarget=Join-Path $folder 'link-target.accdb';Copy-Item -LiteralPath $BuildPath -Destination $linkedTarget
 $db=$app.CurrentDb();$table=$db.CreateTableDef('SyntheticLinkedReports');$table.Connect=';DATABASE='+$linkedTarget;$table.SourceTableName='Report';$db.TableDefs.Append($table)
 Rejected (Join-Path $folder 'linked-out.accdb') 'Linked tables cannot be backed up'
 $db.TableDefs.Delete('SyntheticLinkedReports');$db.Close()
 AssertBackup (-not (Test-Path -LiteralPath (Join-Path $folder 'linked-out.accdb'))) 'Rejected linked-table backup left a file'
 AssertBackup ([int]$app.Eval('DCount("*","MSysObjects","Name Like ''zSwiftBackup*''")') -eq 0) 'Backup left temporary queries'
 #A journal failure after verification keeps the verified backup and its receipt.
 $journalFail=Join-Path $folder 'journal-fail.accdb'
 $db=$app.CurrentDb();$db.TableDefs.Item('OperationJournal').Fields.Item('OperationKind').ValidationRule="<>'Data backup'";$db.Close()
 $outcome=$app.Eval('TryWorkspaceBackup('+(Quote $journalFail)+')')
 $db=$app.CurrentDb();$db.TableDefs.Item('OperationJournal').Fields.Item('OperationKind').ValidationRule='';$db.Close()
 AssertBackup ($outcome -like 'REJECTED:*written and verified*journal entry failed*') ('Journal failure was not reported as a kept backup: '+$outcome)
 $status=Get-Content -LiteralPath ($journalFail+'.verification.json') -Raw -Encoding UTF8|ConvertFrom-Json
 AssertBackup ((Test-Path -LiteralPath $journalFail) -and $status.Status -eq 'PASS' -and $status.SHA256 -eq (Get-FileHash -LiteralPath $journalFail).Hash) 'Journal failure deleted the verified backup or left a mismatched receipt'
 $results.Add('Journal failure after verification keeps the backup and a matching receipt')
 $app.CloseCurrentDatabase();$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
 $status=Get-Content -LiteralPath ($backup+'.verification.json') -Raw -Encoding UTF8|ConvertFrom-Json
 AssertBackup ($status.Status -eq 'PASS' -and $status.SHA256 -eq (Get-FileHash -LiteralPath $backup).Hash -and $status.Destination -eq $backup) 'Verification record does not match the backup file'
 $rows=CompareFiles $backup $source @('OperationJournal','AuditEvent','AuditActor')
 AssertBackup ($rows -ge $status.RowsVerified) 'Independent comparison checked fewer rows than the backup verified'
 $results.Add('Independent exact comparison of '+$rows+' rows; verification record matches the file hash')
 #Restore is the production Settings & Help import into a new blank workspace.
 Copy-Item -LiteralPath $BuildPath -Destination $restored
 $backupHash=(Get-FileHash -LiteralPath $backup).Hash
 $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($restored,$true)
 $app.LoadFromText(5,'WorkspaceImportHarness',(Join-Path $root 'tests/WorkspaceImportHarness.bas'))
 [void]$app.Eval('ConfigureWorkspace("SYNTHETIC RESTORE BOOTSTRAP","Tester")')
 $outcome=$app.Eval('TryWorkspaceImport('+(Quote $backup)+')');AssertBackup ($outcome -like 'PASS:*') ('Restore import failed: '+$outcome)
 $app.CloseCurrentDatabase();$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
 AssertBackup ((Get-FileHash -LiteralPath $backup).Hash -eq $backupHash) 'Restore changed the backup file'
 $db=$engine.OpenDatabase($restored,$false,$true)
 try{
  $r=$db.OpenRecordset('SELECT * FROM RecoveryDraft',4)
  try{
   AssertBackup (-not $r.EOF -and $r.Fields.Item('DraftTitle').Value -ceq 'Pending backup title' -and $r.Fields.Item('StartEntry').Value -ceq 'not-a-date') 'Restore lost raw recovery values'
   foreach($block in 33..39){foreach($entry in 1..3){AssertBackup ($r.Fields.Item('Trait'+$block+'Comment'+$entry).Value -ceq ('BACKUP COMMENT '+$block+'.'+$entry)) 'Restore lost a Chief recovery comment'}}
  }finally{$r.Close()}
  $r=$db.OpenRecordset('SELECT Title FROM Report',4);try{AssertBackup ($r.Fields.Item(0).Value -ceq 'Saved backup title') 'Restore changed committed data'}finally{$r.Close()}
  $r=$db.OpenRecordset('SELECT CommandName FROM Workspace',4);try{AssertBackup ($r.Fields.Item(0).Value -ceq 'SYNTHETIC BACKUP') 'Restore did not bring the backed-up workspace setup'}finally{$r.Close()}
 }finally{$db.Close()}
 $results.Add('Restore through Settings & Help import: raw date, all 21 Chief comments and saved report retained')
 #A backup changed after verification is refused while its receipt sits beside it.
 $changed=Join-Path $folder 'changed-backup.accdb';$changedTarget=Join-Path $folder 'changed-restore.accdb'
 Copy-Item -LiteralPath $backup -Destination $changed;Copy-Item -LiteralPath ($backup+'.verification.json') -Destination ($changed+'.verification.json')
 $db=$engine.OpenDatabase($changed);try{$db.Execute('DELETE FROM RevisionChange')}finally{$db.Close()}
 Copy-Item -LiteralPath $BuildPath -Destination $changedTarget
 $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($changedTarget,$true)
 $app.LoadFromText(5,'WorkspaceImportHarness',(Join-Path $root 'tests/WorkspaceImportHarness.bas'))
 [void]$app.Eval('ConfigureWorkspace("SYNTHETIC CHANGED RESTORE","Tester")')
 $outcome=$app.Eval('TryWorkspaceImport('+(Quote $changed)+')')
 $reports=[int]$app.Eval('DCount("*","Report")')
 $app.CloseCurrentDatabase();$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
 AssertBackup ($outcome -like 'REJECTED:*changed after it was verified*' -and $reports -eq 0) ('Changed backup was restored: '+$outcome)
 $results.Add('A backup changed after verification is refused while its receipt is present')
 if($CompiledPath){
  $compiled=Join-Path $folder 'compiled.accde';$compiledBackup=Join-Path $folder 'compiled-backup.accdb'
  Copy-Item -LiteralPath $CompiledPath -Destination $compiled
  $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($compiled,$true)
  [void]$app.Eval('ConfigureWorkspace("SYNTHETIC COMPILED BACKUP","Tester")')
  $evidence=$app.Eval('WriteWorkspaceBackup('+(Quote $compiledBackup)+')')
  $app.CloseCurrentDatabase();$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
  AssertBackup ($evidence -like '{"Status":"PASS"*') ('Compiled backup failed: '+$evidence)
  [void](CompareFiles $compiledBackup $compiled @('OperationJournal','AuditEvent','AuditActor'))
  $results.Add('ACCDE workspace writes a verified data backup')
 }
}catch{$primary=$_;throw}finally{
 if($app){try{$app.Quit(2)}catch{$cleanup+=$_.Exception.Message};try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+=$_.Exception.Message}}
 foreach($path in $originals.Keys){if((Get-FileHash -LiteralPath $path).Hash -ne $originals[$path]){$cleanup+='Supplied artifact changed'}}
 if($cleanup.Count){if($primary){Write-Warning ($cleanup -join '; ')}else{throw ($cleanup -join '; ')}}
}
@{Status='PASS';InstalledSource=[bool]$InstallSource;OriginalHashes=$originals;Cases=@($results.ToArray());Folder=$folder}|ConvertTo-Json -Depth 4|Set-Content (Join-Path $folder 'verification.json')
Write-Output ('PASS: '+$results.Count+' backup/restore cases and 21 Chief comments; evidence: '+$folder)

}

Invoke-SuiteCase 'persistence' {

$scratch=Join-Path $root ('build/tests/persistence-'+[guid]::NewGuid()+'.accdb')
$before=(Get-FileHash -LiteralPath $BuildPath).Hash
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=New-Object -ComObject Access.Application
try {
    $app.AutomationSecurity=1
    $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'PersistenceHarness',(Join-Path $root 'tests/PersistenceHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestPersistence()')
    Write-Output $result
    if($result -notlike 'PASS:*'){throw $result}
} finally {
    $app.Quit(2)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $before){throw 'Persistence checks changed the supplied build'}
}

}

# VerifyRecovery itself checks that the fresh session has no active contributor,
# displays frmContributorStart and accepts the chosen synthetic contributor.
# Do not call SelectSyntheticContributor (test-contributor-context.ps1) before it:
# that would bypass the startup-selection assertions. Timer checks keep that session.

Invoke-SuiteCase 'recovery' {

$scratch=Join-Path $root ('build\tests\recovery-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'RecoveryHarness',(Join-Path $root 'tests\RecoveryHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $prepared=$app.Eval('PrepareRecovery()');if(-not $prepared.StartsWith('PASS:')){throw $prepared}
    $app.CloseCurrentDatabase();$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
    $preparedCopy=$scratch.Replace('.accdb','-prepared.accdb')
    Copy-Item -LiteralPath $scratch -Destination $preparedCopy
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $verified=$app.Eval('VerifyRecovery()');if(-not $verified.StartsWith('PASS:')){throw $verified}
    if(-not $app.Eval('PrepareTimerCheck()')){throw $app.Eval('RecoveryTimerError()')}
    Start-Sleep -Milliseconds 700
    if(-not $app.Eval('FinishTimerCheck()')){throw $app.Eval('RecoveryTimerError()')}
    Write-Output $prepared;Write-Output $verified
    $app.CloseCurrentDatabase()
    @{Preparation=$prepared;Recovery=$verified;AutomaticTimerAndActiveTextVerified=$true;ReopenedInNewAccessSession=$true;PreparedCheckpointDatabase=$preparedCopy;TestDatabase=$scratch;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash}|ConvertTo-Json|Set-Content (Join-Path $root 'build\tests\recovery-results.json')
}finally{if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}

}

Complete-SuiteRun 'recovery' $BuildPath $buildHash
