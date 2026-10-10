# Additive schema 24: existing report/recovery values and revision history remain intact.
$blockDB=$app.CurrentDb()
foreach($table in @('Report','RecoveryDraft')){
    if(-not @($blockDB.TableDefs.Item($table).Fields|ForEach-Object {$_.Name}).Contains('CommandEmployment')){
        $blockDB.Execute('ALTER TABLE ['+$table+'] ADD COLUMN CommandEmployment LONGTEXT',128)
    }
    $blockDB=$app.CurrentDb()
    $blockDB.TableDefs.Refresh()
    $blockDB.TableDefs.Item($table).Fields.Refresh()
    $blockDB.TableDefs.Item($table).Fields.Item('CommandEmployment').AllowZeroLength=$true
}
$blockDB.Execute('UPDATE AppMetadata SET SchemaVersion=25,AppVersion=''0.25.0-dev''',128)
$app.LoadFromText(5,'modSwift',(Join-Path $root 'src/modSwift.bas'))
$app.LoadFromText(5,'modAdministration',(Join-Path $root 'src/modAdministration.bas'))
$app.LoadFromText(5,'modRecovery',(Join-Path $root 'src/modRecovery.bas'))
$app.LoadFromText(5,'modRevisionCompare',(Join-Path $root 'src/modRevisionCompare.bas'))
$app.LoadFromText(5,'modMajorCore',(Join-Path $root 'src/modMajorCore.bas'))
$app.LoadFromText(5,'modNAVFITAdapter',(Join-Path $root 'src/modNAVFITAdapter.bas'))
$app.LoadFromText(5,'modPDFMapping',(Join-Path $root 'src/modPDFMapping.bas'))
$app.LoadFromText(5,'modNAVFITMapping',(Join-Path $root 'src/modNAVFITMapping.bas'))
$app.DoCmd.OpenForm('frmEditor',1);$f=$app.Forms.Item('frmEditor')
if(-not @($f.Controls|ForEach-Object {$_.Name}).Contains('txtCommandEmployment')){
    [void](Control $f 109 'txtCommandEmployment' 10 10 100 100);$script:swiftControl.Visible=$false
}
SaveScreen $f 'frmEditor'

$workerSource=@('Attribute VB_Name = "modBackupWorker"','Option Explicit','Public Function BackupWorkerScript() As String','    Dim result As String')
foreach($line in (Get-Content (Join-Path $root 'scripts/workspace-backup.ps1'))){$workerSource+='    result = result & "'+$line.Replace('"','""')+'" & vbCrLf'}
$workerSource+=@('    BackupWorkerScript = result','End Function')
$workerPath=Join-Path $root 'build/modBackupWorker.bas';[IO.File]::WriteAllText($workerPath,($workerSource -join "`r`n"))
$app.LoadFromText(5,'modBackupWorker',$workerPath)
