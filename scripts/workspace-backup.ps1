param(
    [Parameter(Mandatory=$true)][ValidateSet('Backup','Restore')][string]$Mode,
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [Parameter(Mandatory=$true)][string]$DestinationPath,
    [ValidateRange(0,60)][int]$WaitForCloseSeconds=0,
    [ValidateRange(0,1000)][int]$MaximumSchemaVersion=0
)
$ErrorActionPreference='Stop'
$SourcePath=[IO.Path]::GetFullPath($SourcePath)
$DestinationPath=[IO.Path]::GetFullPath($DestinationPath)
$extension=[IO.Path]::GetExtension($SourcePath)
if($extension -notin @('.accdb','.accde') -or [IO.Path]::GetExtension($DestinationPath) -ine $extension){throw 'Use .accdb or .accde files with matching source/destination extensions.'}
if(-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)){throw 'Source file does not exist.'}
if($MaximumSchemaVersion -eq 0){
    if(-not $PSScriptRoot){throw 'An embedded backup worker requires its application schema ceiling.'}
    $schemaPath=Join-Path (Split-Path $PSScriptRoot -Parent) 'src/major-schema.json'
    $MaximumSchemaVersion=[int](Get-Content -LiteralPath $schemaPath -Raw|ConvertFrom-Json).SchemaVersion
    if($MaximumSchemaVersion -lt 1){throw 'Canonical backup schema ceiling is unavailable.'}
}
if(Test-Path -LiteralPath $DestinationPath){throw 'Destination already exists. Restore always creates a new file.'}
$evidencePath=$DestinationPath+'.verification.json'
if(Test-Path -LiteralPath $evidencePath){throw 'Destination verification file already exists; it was not overwritten.'}
if(-not (Test-Path -LiteralPath (Split-Path $DestinationPath -Parent))){throw 'Destination folder does not exist.'}
$lockPath=[IO.Path]::ChangeExtension($SourcePath,'.laccdb')
$temporary=Join-Path (Split-Path $DestinationPath -Parent) ('verified-copy-'+[guid]::NewGuid()+'.accdb')
$temporaryEvidence=$temporary+'.verification.json'
$inputStream=$null;$outputStream=$null;$db=$null;$reader=$null;$metadata=$null;$engine=$null;$primary=$null;$cleanup=@()
try {
    # Deny concurrent readers and writers throughout the copy. Never copy an open workspace.
    $deadline=[datetime]::UtcNow.AddSeconds($WaitForCloseSeconds)
    do {
        if(-not (Test-Path -LiteralPath $lockPath)){
            try{$inputStream=[IO.File]::Open($SourcePath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None);break}catch{
                $openFailure=$_.Exception
                while($openFailure.InnerException){$openFailure=$openFailure.InnerException}
                if($openFailure -isnot [IO.IOException] -or ($openFailure.HResult -band 0xFFFF) -notin @(32,33)){throw}
            }
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
    if($schema -lt 1 -or $schema -gt $MaximumSchemaVersion){throw ('Unsupported Swift schema '+$schema+'; maximum supported is '+$MaximumSchemaVersion+'.')}
    $metadata.Close();$metadata=$null;$counts=[ordered]@{}
    foreach($table in $db.TableDefs){
        if($table.Name.StartsWith('MSys')){continue}
        if($table.Attributes -band 0x60000000){throw 'Linked tables cannot be backed up as one file.'}
        $r=$db.OpenRecordset('SELECT Count(*) FROM ['+$table.Name.Replace(']',']]')+']',4)
        try{$counts[$table.Name]=$r.Fields.Item(0).Value}finally{$r.Close()}
    }
    $db.Close();$db=$null;$reader.Close();$reader=$null
    if((Get-FileHash -LiteralPath $temporary -Algorithm SHA256).Hash -cne $sourceHash){throw 'Verification changed copied bytes.'}
    $evidence=[ordered]@{Status='PASS';Mode=$Mode;Source=$SourcePath;Destination=$DestinationPath;SHA256=$sourceHash;Schema=$schema;AppVersion=$version;Counts=$counts;ReopenedReadOnly=$true}
    $evidence|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $temporaryEvidence -Encoding utf8
    # Publish without overwriting either a data file or another verification record.
    [IO.File]::Move($temporary,$DestinationPath)
    [IO.File]::Move($temporaryEvidence,$evidencePath)
    $inputStream.Dispose();$inputStream=$null
    $evidence|ConvertTo-Json -Depth 5
} catch {
    $primary=$_;throw
} finally {
    foreach($owned in @($metadata,$db,$reader)){
        if($null -ne $owned){try{$owned.Close()}catch{$cleanup+='Close: '+$_.Exception.Message}}
    }
    foreach($owned in @($metadata,$db,$reader,$engine)){
        if($null -ne $owned -and [Runtime.InteropServices.Marshal]::IsComObject($owned)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($owned)}catch{$cleanup+='COM release: '+$_.Exception.Message}}
    }
    foreach($owned in @($outputStream,$inputStream)){if($null -ne $owned){try{$owned.Dispose()}catch{$cleanup+='Stream close: '+$_.Exception.Message}}}
    foreach($ownedPath in @($temporary,$temporaryEvidence)){try{if(Test-Path -LiteralPath $ownedPath){Remove-Item -LiteralPath $ownedPath}}catch{$cleanup+='Temporary cleanup: '+$_.Exception.Message}}
    if($cleanup.Count){if($primary){Write-Warning ($cleanup -join '; ')}else{throw ($cleanup -join '; ')}}
}
