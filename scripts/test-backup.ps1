param([Parameter(Mandatory=$true)][string]$BuildPath,[switch]$InstallSource,[string]$CompiledPath)
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
$folder=Join-Path $root ('build/tests/backup-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
$source=Join-Path $folder "O'Brien & `$literal [source].accdb"
$copy=Join-Path $folder 'copy & $literal [verified].accdb'
$restored=Join-Path $folder 'restored.accdb'
$originals=@{};foreach($path in @($BuildPath,$CompiledPath|Where-Object {$_})){$originals[[IO.Path]::GetFullPath($path)]=(Get-FileHash -LiteralPath $path).Hash}
Copy-Item -LiteralPath $BuildPath -Destination $source
$app=$null;$worker=$null;$primary=$null;$cleanup=@();$results=[Collections.Generic.List[string]]::new()
function AssertBackup([bool]$ok,[string]$message){if(-not $ok){throw $message}}
function RejectedCopy([string]$inputPath,[string]$destination,[string]$expected,[switch]$Existing){
 $log=Join-Path $folder ('rejection-'+$results.Count+'.log');$savedPreference=$ErrorActionPreference
 try{$ErrorActionPreference='Continue'; & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'workspace-backup.ps1') -Mode Restore -SourcePath $inputPath -DestinationPath $destination *> $log;$code=$LASTEXITCODE}finally{$ErrorActionPreference=$savedPreference}
 AssertBackup ($code -ne 0) 'Expected backup/restore failure was accepted'
 AssertBackup ((Get-Content -LiteralPath $log -Raw) -match $expected) ('Primary failure was lost: '+$expected)
 if(-not $Existing){AssertBackup (-not (Test-Path -LiteralPath $destination)) 'Rejected copy published a destination'}
 $results.Add('Rejected: '+$expected)
}
try {
 Write-Output ('Controller: backup/recovery runner; disposable database: '+$source)
 $app=New-Object -ComObject Access.Application;$app.AutomationSecurity=1;$app.OpenCurrentDatabase($source,$true)
 if($InstallSource){
  $lines=@('Attribute VB_Name = "modBackupWorker"','Option Explicit','Public Function BackupWorkerScript() As String',' Dim result As String')
  foreach($line in Get-Content -LiteralPath (Join-Path $PSScriptRoot 'workspace-backup.ps1')){$lines+=' result = result & "'+$line.Replace('"','""')+'" & vbCrLf'}
  $lines+=@(' BackupWorkerScript = result','End Function')
  $modulePath=Join-Path $folder 'modBackupWorker.bas';[IO.File]::WriteAllLines($modulePath,$lines,[Text.UTF8Encoding]::new($false))
  $app.LoadFromText(5,'modBackupWorker',$modulePath);$app.LoadFromText(5,'modBackupUI',(Join-Path $root 'src/modBackupUI.bas'))
 }
 $expectedWorker=((Get-Content -LiteralPath (Join-Path $PSScriptRoot 'workspace-backup.ps1')) -join "`r`n")+"`r`n"
 AssertBackup ($app.Eval('BackupWorkerScript()') -ceq $expectedWorker) 'Embedded backup worker differs from source'
 $app.LoadFromText(5,'BackupHarness',(Join-Path $root 'tests/BackupHarness.bas'));$app.DoCmd.RunCommand(126)
 $result=$app.Eval('PrepareBackupCheckpoint()');AssertBackup ($result -like 'PASS:*') $result
 $command=$app.Eval('BackupWorkerCommand("Backup","'+$source+'","'+$copy+'")')
 $protectedResult=Join-Path $folder 'protected-result.accdb'
 $failureCommand=$app.Eval('BackupWorkerCommand("Restore","'+$copy+'","'+$protectedResult+'")')
 $encoded=$command.Substring($command.LastIndexOf(' ')+1)
 $worker=Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile','-NonInteractive','-EncodedCommand',$encoded) -WindowStyle Hidden -PassThru
 Write-Output ('Owned copy worker PID: '+$worker.Id+'; started: '+$worker.StartTime.ToUniversalTime().ToString('o'))
 Start-Sleep -Milliseconds 800
 AssertBackup (-not (Test-Path -LiteralPath $copy)) 'Worker copied an open workspace'
 AssertBackup ([bool]$app.Eval('UXCloseEditorForBackup()')) 'Checkpoint/close for backup failed'
 $app.CloseCurrentDatabase();$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
 AssertBackup ($worker.WaitForExit(15000)) 'Owned copy worker did not finish after normal close'
 if($worker.ExitCode -ne 0){$reason=if(Test-Path -LiteralPath ($copy+'.verification.json')){(Get-Content -LiteralPath ($copy+'.verification.json') -Raw|ConvertFrom-Json).Reason}else{'no verification result'};throw ('Copy worker failed: '+$reason)}
 $status=Get-Content -LiteralPath ($copy+'.verification.json') -Raw|ConvertFrom-Json
 $hash=(Get-FileHash -LiteralPath $source).Hash
 AssertBackup ($status.Status -eq 'PASS' -and $status.SHA256 -eq $hash -and (Get-FileHash -LiteralPath $copy).Hash -eq $hash) 'Verified copy bytes differ'
 & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'workspace-backup.ps1') -Mode Restore -SourcePath $copy -DestinationPath $restored *> (Join-Path $folder 'restore.log')
 AssertBackup ($LASTEXITCODE -eq 0 -and (Get-FileHash -LiteralPath $restored).Hash -eq $hash) 'Restore bytes differ'
 $results.Add('Guided wait-for-close, literal paths, backup/restore and source hashes')
 RejectedCopy $copy $restored 'Destination already exists' -Existing
 AssertBackup ((Get-FileHash -LiteralPath $restored).Hash -eq $hash) 'Rejected overwrite changed existing data'
 RejectedCopy $copy (Join-Path $folder 'missing/out.accdb') 'Destination folder does not exist'
 RejectedCopy (Join-Path $folder 'absent.accdb') (Join-Path $folder 'absent-out.accdb') 'Source file does not exist'
 RejectedCopy $copy (Join-Path $folder 'wrong-extension.accde') 'matching source/destination extensions'
 $evidencePath=$protectedResult+'.verification.json';[IO.File]::WriteAllText($evidencePath,'SYNTHETIC PREEXISTING VERIFICATION')
 $evidenceHash=(Get-FileHash -LiteralPath $evidencePath).Hash
 RejectedCopy $copy $protectedResult 'verification file already exists'
 $encoded=$failureCommand.Substring($failureCommand.LastIndexOf(' ')+1);$savedPreference=$ErrorActionPreference
 try{$ErrorActionPreference='Continue'; & powershell.exe -NoProfile -NonInteractive -EncodedCommand $encoded *> (Join-Path $folder 'guided-sidecar-rejection.log');$code=$LASTEXITCODE}finally{$ErrorActionPreference=$savedPreference}
 AssertBackup ($code -ne 0 -and -not (Test-Path -LiteralPath $protectedResult) -and (Get-FileHash -LiteralPath $evidencePath).Hash -eq $evidenceHash) 'Guided failure replaced preexisting verification evidence'
 $results.Add('Guided primary failure and preexisting verification evidence preserved')
 $hold=[IO.File]::Open($copy,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
 try{RejectedCopy $copy (Join-Path $folder 'locked-out.accdb') 'Close every Access session'}finally{$hold.Dispose()}
 $corrupt=Join-Path $folder 'corrupt.accdb';[IO.File]::WriteAllText($corrupt,'SYNTHETIC NOT AN ACCESS DATABASE')
 RejectedCopy $corrupt (Join-Path $folder 'corrupt-out.accdb') 'Unrecognized database format|not a valid|not.*database'
 $future=Join-Path $folder 'future.accdb';Copy-Item -LiteralPath $copy -Destination $future
 $engine=New-Object -ComObject DAO.DBEngine.120;$db=$engine.OpenDatabase($future)
 try{$maximum=[int](Get-Content (Join-Path $root 'src/major-schema.json') -Raw|ConvertFrom-Json).SchemaVersion;$db.Execute(('UPDATE AppMetadata SET SchemaVersion='+($maximum+1)),128)}finally{$db.Close()}
 $futureHash=(Get-FileHash -LiteralPath $future).Hash
 RejectedCopy $future (Join-Path $folder 'future-out.accdb') 'Unsupported Swift schema'
 AssertBackup ((Get-FileHash -LiteralPath $future).Hash -eq $futureHash) 'Future-schema rejection changed its source'
 $linked=Join-Path $folder 'linked.accdb';Copy-Item -LiteralPath $copy -Destination $linked
 $db=$engine.OpenDatabase($linked)
 try{$table=$db.CreateTableDef('SyntheticLinkedReports');$table.Connect=';DATABASE='+$copy;$table.SourceTableName='Report';$db.TableDefs.Append($table)}finally{$db.Close()}
 $linkedHash=(Get-FileHash -LiteralPath $linked).Hash
 RejectedCopy $linked (Join-Path $folder 'linked-out.accdb') 'Linked tables cannot be backed up as one file'
 AssertBackup ((Get-FileHash -LiteralPath $linked).Hash -eq $linkedHash -and (Get-FileHash -LiteralPath $copy).Hash -eq $hash) 'Linked-table rejection changed its source or linked target'
 $db=$engine.OpenDatabase($restored,$true,$true)
 try{
  $r=$db.OpenRecordset('SELECT * FROM RecoveryDraft',4)
  try{
   AssertBackup (-not $r.EOF -and $r.Fields.Item('DraftTitle').Value -ceq 'Pending backup title' -and $r.Fields.Item('StartEntry').Value -ceq 'not-a-date') 'Restore lost raw recovery values'
   foreach($block in 33..39){foreach($entry in 1..3){AssertBackup ($r.Fields.Item('Trait'+$block+'Comment'+$entry).Value -ceq ('BACKUP COMMENT '+$block+'.'+$entry)) 'Restore lost a Chief recovery comment'}}
  }finally{$r.Close()}
  $r=$db.OpenRecordset('SELECT Title FROM Report',4);try{AssertBackup ($r.Fields.Item(0).Value -ceq 'Saved backup title') 'Restore changed committed data'}finally{$r.Close()}
 }finally{$db.Close()}
 $results.Add('Raw date and all 21 Chief comments retained independently of saved report')
 if($CompiledPath){
  $compiledCopy=Join-Path $folder 'compiled-copy.accde'
  & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'workspace-backup.ps1') -Mode Backup -SourcePath $CompiledPath -DestinationPath $compiledCopy *> (Join-Path $folder 'compiled.log')
  AssertBackup ($LASTEXITCODE -eq 0 -and (Get-FileHash -LiteralPath $compiledCopy).Hash -eq $originals[[IO.Path]::GetFullPath($CompiledPath)]) 'Compiled backup bytes differ'
  $compiledRestore=Join-Path $folder 'compiled-restored.accde'
  & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'workspace-backup.ps1') -Mode Restore -SourcePath $compiledCopy -DestinationPath $compiledRestore *> (Join-Path $folder 'compiled-restore.log')
  AssertBackup ($LASTEXITCODE -eq 0 -and (Get-FileHash -LiteralPath $compiledRestore).Hash -eq $originals[[IO.Path]::GetFullPath($CompiledPath)]) 'Compiled restore bytes differ'
  $db=$engine.OpenDatabase($compiledRestore,$true,$true);try{AssertBackup ($db.Properties.Item('MDE').Value -eq 'T') 'Compiled restore lost its compiled identity'}finally{$db.Close()}
  RejectedCopy $compiledCopy (Join-Path $folder 'compiled-as-accdb.accdb') 'matching source/destination extensions'
  $results.Add('ACCDE byte-preserving backup/restore and extension protection')
 }
 AssertBackup (@(Get-ChildItem -LiteralPath $folder -Filter 'verified-copy-*').Count -eq 0) 'Rejected copies left temporary output files'
}catch{$primary=$_;throw}finally{
 if($app){try{$app.Quit(2)}catch{$cleanup+=$_.Exception.Message};try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}catch{$cleanup+=$_.Exception.Message}}
 if($worker){try{if(-not $worker.HasExited){$worker.Kill();[void]$worker.WaitForExit(5000)}}catch{$cleanup+='Owned worker cleanup: '+$_.Exception.Message}}
 foreach($path in $originals.Keys){if((Get-FileHash -LiteralPath $path).Hash -ne $originals[$path]){$cleanup+='Supplied artifact changed'}}
 if($cleanup.Count){if($primary){Write-Warning ($cleanup -join '; ')}else{throw ($cleanup -join '; ')}}
}
@{Status='PASS';InstalledSource=[bool]$InstallSource;OriginalHashes=$originals;Cases=@($results.ToArray());Folder=$folder}|ConvertTo-Json -Depth 4|Set-Content (Join-Path $folder 'verification.json')
Write-Output ('PASS: '+$results.Count+' backup/restore fault cases and 21 Chief comments; evidence: '+$folder)
