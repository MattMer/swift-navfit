param([Parameter(Mandatory=$true)][string]$BuildPath)
#Front end and data file (S01): a blank copy writes a data file, links to it, saves
#through the links, keeps the link in a new session, reports a missing data file and
#works again once pointed at the moved file.
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$folder=Join-Path $root ('build/tests/data-link-'+[guid]::NewGuid());[void](New-Item -ItemType Directory -Path $folder)
$front=Join-Path $folder 'front end.accdb';$data=Join-Path $folder 'workspace data.accdb';$moved=Join-Path $folder 'moved workspace data.accdb'
$other=Join-Path $folder 'not a data file.accdb';$backup=Join-Path $folder 'linked backup.accdb'
Copy-Item -LiteralPath $BuildPath -Destination $front
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$engine=New-Object -ComObject DAO.DBEngine.120
$empty=$engine.CreateDatabase($other,';LANGID=0x0409;CP=1252;COUNTRY=0',128);$empty.Close()

#Tables, fields, indexes, rules and relationships as text. Access names an unnamed
#primary key at random, so those names are left out.
function Get-SchemaText([string]$Path){
    $db=$engine.OpenDatabase($Path,$false,$true);$lines=New-Object Collections.Generic.List[string]
    try{
        foreach($t in ($db.TableDefs|Where-Object {$_.Name -notlike 'MSys*' -and $_.Name -notlike '~*'}|Sort-Object Name)){
            $lines.Add('TABLE '+$t.Name+' rule=['+[string]$t.ValidationRule+'] rows='+$t.RecordCount)
            foreach($f in $t.Fields){$lines.Add($t.Name+' F '+$f.Name+' type='+$f.Type+' size='+$f.Size+' req='+$f.Required+' attr='+$f.Attributes+' zero='+$f.AllowZeroLength+' default=['+[string]$f.DefaultValue+']')}
            foreach($i in $t.Indexes){$lines.Add($t.Name+' I '+($i.Name -replace '^Index_[0-9A-F_]+$','Index_AUTO')+' unique='+$i.Unique+' primary='+$i.Primary+' foreign='+$i.Foreign+' fields='+(@($i.Fields|ForEach-Object {$_.Name}) -join ','))}
        }
        foreach($r in ($db.Relations|Where-Object {$_.Table -notlike 'MSys*'})){$lines.Add('REL '+$r.Name+' '+$r.Table+'->'+$r.ForeignTable+' attr='+$r.Attributes+' '+(@($r.Fields|ForEach-Object {$_.Name+'='+$_.ForeignName}) -join ','))}
        $code=0;foreach($kind in 'Modules','Forms','Reports','Scripts'){$code+=$db.Containers.Item($kind).Documents.Count}
    }finally{$db.Close()}
    [pscustomobject]@{Lines=@($lines|Sort-Object);CodeObjects=$code}
}
function Start-Front{
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($front,$true)
    $app.LoadFromText(5,'DataLinkHarness',(Join-Path $root 'tests/DataLinkHarness.bas'))
    $app.DoCmd.RunCommand(126)
    if(-not $app.IsCompiled){throw 'VBA did not compile after loading DataLinkHarness'}
    $app
}
function Invoke-Check($app,[string]$Expression){
    $result=[string]$app.Eval($Expression)
    if($result -notlike 'PASS:*'){throw $result}
    Write-Output $result
}
function Stop-Front($app){if($null -ne $app){$app.Quit(1);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}

$app=$null
try{
    $app=Start-Front
    foreach($name in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
    Invoke-Check $app ('TestDataFileCreate("'+$data+'")')
    $source=Get-SchemaText $BuildPath;$created=Get-SchemaText $data
    $difference=@(Compare-Object $source.Lines $created.Lines)
    if($difference.Count){throw ('Data file schema differs from the build: '+(($difference|Select-Object -First 3|ForEach-Object {$_.SideIndicator+' '+$_.InputObject}) -join ' | '))}
    if($created.CodeObjects -ne 0){throw ('Data file holds '+$created.CodeObjects+' code objects')}
    Write-Output ('PASS: data file matches the build: '+$source.Lines.Count+' schema lines, no modules, forms, reports or macros')
    Invoke-Check $app ('TestDataFileLink("'+$data+'","'+$other+'")')
    Invoke-Check $app ('TestDataFileBackup("'+$backup+'")')
}finally{Stop-Front $app;$app=$null}

#The saved report is in the data file, and the front end holds no table of its own.
$db=$engine.OpenDatabase($data,$false,$true)
try{$rows=$db.OpenRecordset('SELECT Count(*) FROM Report',4);$saved=[int]$rows.Fields.Item(0).Value;$rows.Close()}finally{$db.Close()}
if($saved -ne 1){throw ('Data file holds '+$saved+' reports, expected 1')}
$db=$engine.OpenDatabase($front,$false,$true)
try{$local=@($db.TableDefs|Where-Object {$_.Name -notlike 'MSys*' -and $_.Name -notlike '~*' -and -not $_.Connect}).Count}finally{$db.Close()}
if($local -ne 0){throw ('Front end still holds '+$local+' local tables')}
Write-Output 'PASS: the saved report is in the data file and the front end holds no tables'

#Restore: the backup imports into a new blank copy through the production import.
$restore=Join-Path $folder 'restore target.accdb';Copy-Item -LiteralPath $BuildPath -Destination $restore
$app=New-Object -ComObject Access.Application
try{
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($restore,$true)
    foreach($name in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
    $app.LoadFromText(5,'WorkspaceImportHarness',(Join-Path $root 'tests/WorkspaceImportHarness.bas'))
    $app.DoCmd.RunCommand(126)
    if(-not $app.IsCompiled){throw 'VBA did not compile after loading WorkspaceImportHarness'}
    $outcome=[string]$app.Eval('TryWorkspaceImport("'+$backup+'")')
    if($outcome -notlike 'PASS:*'){throw ('Restore of the linked backup failed: '+$outcome)}
    if($app.Eval('DCount("*","Report")') -ne 1 -or $app.Eval('DLookup("CommandName","Workspace")') -ne 'SYNTHETIC DATA FILE COMMAND'){throw 'Restored workspace does not hold the saved data'}
}finally{$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null}
Write-Output 'PASS: a backup written from a linked copy restores through the import'

#Move the data file while the application is closed, then start a new session.
Move-Item -LiteralPath $data -Destination $moved
try{
    $app=Start-Front
    if($app.Forms.Count -ne 0){throw 'Startup opened a screen without its data file'}
    Invoke-Check $app ('TestDataFileMoved("'+$data+'","'+$moved+'")')
    if(-not $app.Eval('SwiftAppStartup()')){throw 'Startup failed after relinking'}
    if($app.Forms.Count -eq 0){throw 'Startup opened no screen after relinking'}
}finally{Stop-Front $app;$app=$null}
#From the setup screen, as a user does it: with the setup screen open, a new data file
#is written and linked, and setup opens again on it.
$front=Join-Path $folder 'second front end.accdb';$second=Join-Path $folder 'second data.accdb'
Copy-Item -LiteralPath $BuildPath -Destination $front
try{
    $app=Start-Front
    if(-not $app.CurrentProject.AllForms.Item('frmSetup').IsLoaded){throw 'Setup screen was not open on a blank copy'}
    if(-not $app.Eval('UXUseDataFile("'+$second+'",True)')){throw 'Setup screen data file action failed'}
    if(-not $app.Eval('SwiftDataIsLinked()') -or $app.Eval('SwiftDataPath()') -ne $second){throw 'Setup screen action did not link the new data file'}
    if(-not $app.CurrentProject.AllForms.Item('frmSetup').IsLoaded){throw 'Setup did not reopen on the new data file'}
    if(-not $app.Eval('SaveSetupValues("SYNTHETIC SECOND COMMAND","Tester")')){throw 'Setup failed on the new data file'}
}finally{Stop-Front $app;$app=$null}
Write-Output 'PASS: the setup screen writes a new data file, links it and sets up on it'

#Import from a previous version into a copy that uses a new data file: the rows land in
#the data file and the copy stays linked.
$front=Join-Path $folder 'import front end.accdb';$imported=Join-Path $folder 'import data.accdb'
Copy-Item -LiteralPath $BuildPath -Destination $front
try{
    $app=Start-Front
    $app.LoadFromText(5,'WorkspaceImportHarness',(Join-Path $root 'tests/WorkspaceImportHarness.bas'))
    $app.DoCmd.RunCommand(126)
    if(-not $app.IsCompiled){throw 'VBA did not compile after loading WorkspaceImportHarness'}
    if(-not $app.Eval('UXUseDataFile("'+$imported+'",True)')){throw 'Could not start on a new data file before importing'}
    foreach($name in @($app.CurrentProject.AllForms|Where-Object {$_.IsLoaded}|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$name,2)}
    $outcome=[string]$app.Eval('TryWorkspaceImport("'+$backup+'")')
    if($outcome -notlike 'PASS:*'){throw ('Import into a linked copy failed: '+$outcome)}
    if(-not $app.Eval('SwiftDataIsLinked()') -or $app.Eval('DCount("*","Report")') -ne 1){throw 'Import left the copy unlinked or without the report'}
}finally{Stop-Front $app;$app=$null}
$db=$engine.OpenDatabase($imported,$false,$true)
try{$rows=$db.OpenRecordset('SELECT Count(*) FROM Report',4);$landed=[int]$rows.Fields.Item(0).Value;$rows.Close()}finally{$db.Close()}
if($landed -ne 1){throw ('Imported report is not in the data file: '+$landed)}
Write-Output 'PASS: an import into a linked copy lands in the data file'

#Two sessions on one data file: each front end is its own file, the data file is shared,
#and the other session is counted from the data file.
$front=Join-Path $folder 'front end.accdb';$twin=Join-Path $folder 'twin front end.accdb'
Copy-Item -LiteralPath $front -Destination $twin
$first=$null;$other=$null
try{
    $first=Start-Front
    #Loading the harness reset VBA, which released the hold startup took on the data file.
    [void]$first.Eval('SwiftDataHold()')
    Invoke-Check $first 'TestDataFileSessions(0)'
    $other=New-Object -ComObject Access.Application
    $other.AutomationSecurity=1;$other.OpenCurrentDatabase($twin,$true)
    if(-not $other.Eval('SwiftDataIsLinked()')){throw 'Twin front end is not linked'}
    #The session that starts second finds the data file held, and is warned.
    if(-not $other.Eval('SwiftDataOtherSession()')){throw 'The second session did not see the first on the data file'}
    $notice=[string]$other.Eval('WorkspaceSafetyNotice()')
    if($notice -notlike '*already open in another session*'){throw ('Second session was not warned: '+$notice)}
    Write-Output 'PASS: a second session on the same data file is warned at startup'
}finally{
    if($null -ne $other){$other.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($other)}
    Stop-Front $first
}
[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)
if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $buildHash){throw 'The build was changed'}
Write-Output 'PASS: a moved data file is reported at startup and works again once chosen'
