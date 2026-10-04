param(
    [Parameter(Mandatory=$true)][ValidateSet('Backup','Restore')][string]$Mode,
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [Parameter(Mandatory=$true)][string]$DestinationPath,
    [ValidateRange(0,60)][int]$WaitForCloseSeconds=0
)
$ErrorActionPreference='Stop'
$SourcePath=[IO.Path]::GetFullPath($SourcePath)
$DestinationPath=[IO.Path]::GetFullPath($DestinationPath)
if([IO.Path]::GetExtension($SourcePath) -ine '.accdb' -or [IO.Path]::GetExtension($DestinationPath) -ine '.accdb'){throw 'Use .accdb files.'}
if(Test-Path -LiteralPath $DestinationPath){throw 'Destination already exists. Restore always creates a new file.'}
if(-not (Test-Path -LiteralPath (Split-Path $DestinationPath -Parent))){throw 'Destination folder does not exist.'}
$lockPath=[IO.Path]::ChangeExtension($SourcePath,'.laccdb')
$temporary=Join-Path (Split-Path $DestinationPath -Parent) ('verified-copy-'+[guid]::NewGuid()+'.accdb')
$inputStream=$null;$outputStream=$null;$db=$null;$reader=$null
try {
    # Deny concurrent readers and writers throughout the copy. Never copy an open workspace.
    $deadline=[datetime]::UtcNow.AddSeconds($WaitForCloseSeconds)
    do {
        if(-not (Test-Path -LiteralPath $lockPath)){
            try{$inputStream=[IO.File]::Open($SourcePath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None);break}catch{}
        }
        if([datetime]::UtcNow -ge $deadline){throw 'Close every Access session using the source file before continuing.'}
        Start-Sleep -Milliseconds 250
    } while($true)
    $sha=[Security.Cryptography.SHA256]::Create()
    try{$sourceHash=[BitConverter]::ToString($sha.ComputeHash($inputStream)).Replace('-','')}finally{$sha.Dispose()}
    $inputStream.Position=0
    $outputStream=[IO.File]::Open($temporary,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
    $inputStream.CopyTo($outputStream);$outputStream.Flush($true);$outputStream.Dispose();$outputStream=$null
    $copyHash=(Get-FileHash -LiteralPath $temporary -Algorithm SHA256).Hash
    if($copyHash -cne $sourceHash){throw 'Copied bytes do not match source.'}
    $engine=New-Object -ComObject DAO.DBEngine.120
    $reader=$engine.CreateWorkspace('SwiftBackupVerify','admin','',2)
    $db=$reader.OpenDatabase($temporary,$true,$true)
    $metadata=$db.OpenRecordset('SELECT SchemaVersion,AppVersion FROM AppMetadata',4)
    if($metadata.EOF){throw 'Missing Swift metadata.'}
    $schema=[int]$metadata.Fields.Item('SchemaVersion').Value;$version=[string]$metadata.Fields.Item('AppVersion').Value
    if($schema -lt 1 -or $schema -gt 20){throw 'Unsupported Swift schema.'}
    $metadata.Close();$counts=[ordered]@{}
    foreach($table in $db.TableDefs){
        if($table.Name.StartsWith('MSys')){continue}
        if($table.Attributes -band 0x60000000){throw 'Linked tables cannot be backed up as one file.'}
        $r=$db.OpenRecordset('SELECT Count(*) FROM ['+$table.Name.Replace(']',']]')+']',4)
        try{$counts[$table.Name]=$r.Fields.Item(0).Value}finally{$r.Close()}
    }
    $db.Close();$db=$null;$reader.Close();$reader=$null
    if((Get-FileHash -LiteralPath $temporary -Algorithm SHA256).Hash -cne $sourceHash){throw 'Verification changed copied bytes.'}
    # Move without overwrite only after byte and database verification.
    [IO.File]::Move($temporary,$DestinationPath)
    $inputStream.Dispose();$inputStream=$null
    $evidence=[ordered]@{Status='PASS';Mode=$Mode;Source=$SourcePath;Destination=$DestinationPath;SHA256=$sourceHash;Schema=$schema;AppVersion=$version;Counts=$counts;ReopenedReadOnly=$true}
    $evidence|ConvertTo-Json -Depth 5|Set-Content -LiteralPath ($DestinationPath+'.verification.json') -Encoding utf8
    $evidence|ConvertTo-Json -Depth 5
} finally {
    if($db){$db.Close()}
    if($reader){$reader.Close()}
    if($outputStream){$outputStream.Dispose()}
    if($inputStream){$inputStream.Dispose()}
    if(Test-Path -LiteralPath $temporary){Remove-Item -LiteralPath $temporary}
}
