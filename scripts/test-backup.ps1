param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$token=[guid]::NewGuid().ToString()
$source=Join-Path $root ("build/tests/backup-O'Brien-"+$token+'.accdb')
$copy=Join-Path $root ('build/tests/copy & $literal-'+$token+'.accdb')
$restored=Join-Path $root ('build/tests/restored-'+$token+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $source
$app=$null;$worker=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($source,$true)
    $app.LoadFromText(5,'BackupHarness',(Join-Path $root 'tests/BackupHarness.bas'))
    $result=$app.Eval('PrepareBackupCheckpoint()');if(-not $result.StartsWith('PASS:')){throw $result}
    $command=$app.Eval('BackupWorkerCommand("Backup","'+$source+'","'+$copy+'")')
    $encoded=$command.Substring($command.LastIndexOf(' ')+1)
    $worker=Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile','-NonInteractive','-EncodedCommand',$encoded) -WindowStyle Hidden -PassThru
    Start-Sleep -Milliseconds 800
    if(Test-Path -LiteralPath $copy){throw 'Worker copied an open workspace.'}
    if(-not $app.Eval('UXCloseEditorForBackup()')){throw 'Approved backup close failed.'}
    $app.CloseCurrentDatabase();$app.Quit(2);$app=$null
    if(-not $worker.WaitForExit(15000)){throw 'Copy worker did not finish after closing source.'}
    if($worker.ExitCode -ne 0){throw 'Copy worker failed.'}
    $status=Get-Content -LiteralPath ($copy+'.verification.json') -Raw | ConvertFrom-Json
    if($status.Status -ne 'PASS'){throw 'Copy verification failed.'}
    $hash=(Get-FileHash -LiteralPath $source).Hash
    if($hash -cne (Get-FileHash -LiteralPath $copy).Hash -or $status.SHA256 -cne $hash){throw 'Copy bytes differ.'}
    & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'workspace-backup.ps1') -Mode Restore -SourcePath $copy -DestinationPath $restored *> (Join-Path $root 'build/tests/backup-restore.log')
    if($LASTEXITCODE -ne 0 -or (Get-FileHash -LiteralPath $restored).Hash -cne $hash){throw 'Restore failed byte equality.'}
    $navPreviousErrorPreference=$ErrorActionPreference
    try {
        $ErrorActionPreference='Continue'
        & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'workspace-backup.ps1') -Mode Restore -SourcePath $copy -DestinationPath $restored *> (Join-Path $root 'build/tests/backup-overwrite.log')
        $navOverwriteExitCode=$LASTEXITCODE
    } finally {$ErrorActionPreference=$navPreviousErrorPreference}
    if($navOverwriteExitCode -eq 0 -or (Get-FileHash -LiteralPath $restored).Hash -cne $hash){throw 'Existing destination was not protected.'}
    $engine=New-Object -ComObject DAO.DBEngine.120;$db=$engine.OpenDatabase($restored,$true,$true)
    try {
        $r=$db.OpenRecordset('SELECT DraftTitle,StartEntry FROM RecoveryDraft',4)
        try{if($r.EOF -or $r.Fields.Item(0).Value -cne 'Pending backup title' -or $r.Fields.Item(1).Value -cne 'not-a-date'){throw 'Pending recovery values were lost.'}}finally{$r.Close()}
        $r=$db.OpenRecordset('SELECT Title FROM Report',4)
        try{if($r.Fields.Item(0).Value -cne 'Saved backup title'){throw 'Restore changed committed value.'}}finally{$r.Close()}
    }finally{$db.Close()}
    Write-Output 'PASS: guided copy waits for close; quoted paths; byte hashes; restore; no overwrite; raw checkpoints preserved'
} finally {
    if($app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    if($worker -and -not $worker.HasExited){$worker.Kill()}
}
