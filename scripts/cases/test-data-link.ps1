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

#Tables, fields, indexes, rules and relationships as text. Primary key names are left
#out: a build leaves them for Access to name at random.
function Get-SchemaText([string]$Path,[switch]$NoRows){
    $db=$engine.OpenDatabase($Path,$false,$true);$lines=New-Object Collections.Generic.List[string]
    try{
        foreach($t in ($db.TableDefs|Where-Object {$_.Name -notlike 'MSys*' -and $_.Name -notlike '~*'}|Sort-Object Name)){
            $lines.Add('TABLE '+$t.Name+' rule=['+[string]$t.ValidationRule+']'+$(if($NoRows){''}else{' rows='+$t.RecordCount}))
            foreach($f in $t.Fields){$lines.Add($t.Name+' F '+$f.Name+' type='+$f.Type+' size='+$f.Size+' req='+$f.Required+' attr='+$f.Attributes+' zero='+$f.AllowZeroLength+' default=['+[string]$f.DefaultValue+']')}
            foreach($i in $t.Indexes){$lines.Add($t.Name+' I '+$(if($i.Primary){'PRIMARY'}else{$i.Name})+' unique='+$i.Unique+' primary='+$i.Primary+' foreign='+$i.Foreign+' fields='+(@($i.Fields|ForEach-Object {$_.Name}) -join ','))}
        }
        foreach($r in ($db.Relations|Where-Object {$_.Table -notlike 'MSys*'})){$lines.Add('REL '+$r.Name+' '+$r.Table+'->'+$r.ForeignTable+' attr='+$r.Attributes+' '+(@($r.Fields|ForEach-Object {$_.Name+'='+$_.ForeignName}) -join ','))}
        #A file made by DAO alone, like a schema fixture, has no application containers.
        $code=0;foreach($container in $db.Containers){if($container.Name -in 'Modules','Forms','Reports','Scripts'){$code+=$container.Documents.Count}}
    }finally{$db.Close()}
    [pscustomobject]@{Lines=@($lines|Sort-Object);CodeObjects=$code}
}
function Start-Front{
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($front,$true)
    $app.LoadFromText(5,'DataLinkHarness',(Join-Path $root 'tests/DataLinkHarness.bas'))
    $app.DoCmd.RunCommand(126)
    #A session that will not be handed back is closed here, or it would stay open on its own.
    if(-not $app.IsCompiled){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);throw 'VBA did not compile after loading DataLinkHarness'}
    $app
}
function Invoke-Check($app,[string]$Expression){
    $result=[string]$app.Eval($Expression)
    if($result -notlike 'PASS:*'){throw $result}
    Write-Output $result
}
function Stop-Front($app){if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}

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
try{
    $local=@($db.TableDefs|Where-Object {$_.Name -notlike 'MSys*' -and $_.Name -notlike '~*' -and $_.Name -notlike 'USys*' -and -not $_.Connect}).Count
    $blank=@($db.TableDefs|Where-Object {$_.Name -like 'USysBlank*' -and -not $_.Connect -and $_.RecordCount -ge 0}).Count
}finally{$db.Close()}
if($local -ne 0){throw ('Front end still holds '+$local+' workspace tables')}
if($blank -ne $source.Lines.Where({$_ -like 'TABLE *'}).Count){throw ('Front end keeps '+$blank+' blank tables')}
Write-Output 'PASS: the saved report is in the data file; the front end holds only its hidden blank tables'

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
#The window offered when the data file is missing, and its Create a new data file path,
#on a copy of that front end whose data file is gone.
$lost=Join-Path $folder 'lost front end.accdb';Copy-Item -LiteralPath $front -Destination $lost
$keepFront=$front;$front=$lost;$fresh=Join-Path $folder 'fresh data.accdb'
#Link the copy to a spare data file, then delete that file.
$gone=Join-Path $folder 'gone data.accdb';Copy-Item -LiteralPath $moved -Destination $gone
try{$app=Start-Front;if(-not $app.Eval('LinkDataFile("'+$gone+'")')){throw 'Could not link to the spare data file'}}finally{Stop-Front $app;$app=$null}
Start-Sleep 2;Remove-Item -LiteralPath $gone
try{
    $app=Start-Front
    Invoke-Check $app ('TestDataFileWindow("'+$fresh+'")')
}finally{Stop-Front $app;$app=$null}
$front=$keepFront
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

#The default: Create workspace on the setup screen of a new copy makes the data file beside
#the application and sets up in it. Then the two files are moved together to another folder:
#the copy finds its data file beside it and starts without asking.
$defaultFolder=Join-Path $folder 'default';[void](New-Item -ItemType Directory -Path $defaultFolder)
$front=Join-Path $defaultFolder 'Swift NAVFIT.accdb';Copy-Item -LiteralPath $BuildPath -Destination $front
try{
    $app=Start-Front
    Invoke-Check $app 'TestSetupDefault("SYNTHETIC DEFAULT COMMAND","Tester")'
}finally{Stop-Front $app;$app=$null}
$together=Join-Path $folder 'moved together';[void](New-Item -ItemType Directory -Path $together)
Move-Item -LiteralPath $front -Destination $together;Move-Item -LiteralPath (Join-Path $defaultFolder 'Swift NAVFIT data.accdb') -Destination $together
$front=Join-Path $together 'Swift NAVFIT.accdb'
try{
    $app=Start-Front
    if($app.Eval('SwiftDataPath()') -ne (Join-Path $together 'Swift NAVFIT data.accdb')){throw ('The copy did not find its data file beside it: '+$app.Eval('SwiftDataPath()'))}
    if(-not $app.CurrentProject.AllForms.Item('frmContributorStart').IsLoaded){throw 'Startup did not reach contributor choice after the move'}
    if($app.Eval('DLookup("CommandName","Workspace")') -ne 'SYNTHETIC DEFAULT COMMAND'){throw 'The moved data file does not hold the workspace'}
}finally{Stop-Front $app;$app=$null}
Write-Output 'PASS: setup makes the data file beside the application by default, and the pair still opens after being moved together'

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

#An earlier single-file workspace, opened from the setup screen of a new copy, is brought
#into a new data file beside it with the verified import. The earlier file is only read.
$front=Join-Path $folder 'bring in front end.accdb';Copy-Item -LiteralPath $BuildPath -Destination $front
$restoreHash=(Get-FileHash -LiteralPath $restore).Hash
$broughtIn=Join-Path $folder 'restore target data.accdb'
try{
    $app=Start-Front
    if(-not $app.Eval('IsEarlierWorkspace("'+$restore+'")')){throw 'An earlier workspace was not recognised'}
    if($app.Eval('IsEarlierWorkspace("'+$backup+'")')){throw 'A data file was taken for an earlier workspace'}
    if(-not $app.Eval('UXBringInWorkspace("'+$restore+'",False)')){throw 'Bringing in the earlier workspace failed'}
    if(-not $app.Eval('SwiftDataIsLinked()') -or $app.Eval('SwiftDataPath()') -ne $broughtIn){throw ('Not linked to the new data file beside the workspace: '+$app.Eval('SwiftDataPath()'))}
    if($app.Eval('DCount("*","Report")') -ne 1 -or $app.Eval('DLookup("CommandName","Workspace")') -ne 'SYNTHETIC DATA FILE COMMAND'){throw 'The workspace was not brought in'}
    if(-not $app.CurrentProject.AllForms.Item('frmContributorStart').IsLoaded){throw 'Startup did not reach contributor choice on the brought-in workspace'}
    $refused=[string]$app.Eval('DataFileProblem("'+$restore+'")')
    if($refused -notlike 'That file is a complete Swift NAVFIT workspace*'){throw ('A complete workspace was offered as a data file: '+$refused)}
}finally{Stop-Front $app;$app=$null}
if((Get-FileHash -LiteralPath $restore).Hash -ne $restoreHash){throw 'The earlier workspace was changed'}
$db=$engine.OpenDatabase($broughtIn,$false,$true)
try{$rows=$db.OpenRecordset('SELECT Count(*) FROM Report',4);$landed=[int]$rows.Fields.Item(0).Value;$rows.Close()}finally{$db.Close()}
if($landed -ne 1){throw 'The brought-in report is not in the new data file'}
Write-Output 'PASS: an earlier workspace opened from the setup screen is brought into a new data file beside it, unchanged itself'

#Schema upgrade (S01 step 4). A data file of the previous schema, built from its captured
#definition, upgraded by this version has exactly the structure of this build.
function New-SchemaFixture([int]$Schema,[string]$Path){
    $definition=Join-Path $root ('tests/historical-schemas/schema-'+$Schema+'.sql')
    if(-not (Test-Path -LiteralPath $definition)){throw ('No captured definition for schema '+$Schema+'; capture it with scripts/capture-schema-definition.ps1 before changing the schema')}
    $sql=@();$zeroLengthOff=@();$rules=@{}
    foreach($line in [IO.File]::ReadAllLines($definition,[Text.Encoding]::UTF8)){
        if($line -match '^-- (\w+): (.*)$'){
            if($Matches[1] -eq 'zeroLengthOff'){$zeroLengthOff+=$Matches[2]}
            elseif($Matches[1] -eq 'validationRule'){$p=$Matches[2].Split(' ',2);$rules[$p[0]]=$p[1]}
        }elseif($line.Trim()){$sql+=$line}
    }
    $db=$engine.CreateDatabase($Path,';LANGID=0x0409;CP=1252;COUNTRY=0',128)
    try{
        foreach($statement in $sql){$db.Execute($statement,128)}
        foreach($t in @($db.TableDefs|Where-Object {$_.Name -notlike 'MSys*'})){foreach($f in @($t.Fields)){if($f.Type -in 10,12 -and -not $f.Required){$f.AllowZeroLength=$true}}}
        foreach($item in $zeroLengthOff){if($item){$p=$item.Split('.');$db.TableDefs.Item($p[0]).Fields.Item($p[1]).AllowZeroLength=$false}}
        foreach($rule in $rules.GetEnumerator()){$db.TableDefs.Item($rule.Key).ValidationRule=$rule.Value}
        if(-not @($sql|Where-Object {$_ -like 'INSERT INTO AppMetadata*'}).Count){$db.Execute("INSERT INTO AppMetadata (MetadataID,AppVersion,SchemaVersion) VALUES (1,'captured',"+$Schema+")",128)}
    }finally{$db.Close()}
}
function Invoke-Dao([string]$Path,[string[]]$Sql){$db=$engine.OpenDatabase($Path);try{foreach($statement in $Sql){$db.Execute($statement,128)}}finally{$db.Close()}}
function Get-Schema([string]$Path){$db=$engine.OpenDatabase($Path,$false,$true);try{$r=$db.OpenRecordset('SELECT SchemaVersion FROM AppMetadata',4);[int]$r.Fields.Item(0).Value}finally{$db.Close()}}
$current=Get-Schema $BuildPath;$previous=$current-1
$front=Join-Path $folder 'upgrade runner.accdb';Copy-Item -LiteralPath $BuildPath -Destination $front
$old=Join-Path $folder ('schema '+$previous+' data.accdb');New-SchemaFixture $previous $old;$oldHash=(Get-FileHash -LiteralPath $old).Hash
$app=$null;$holder=$null
try{
    $app=Start-Front
    $result=[string]$app.Eval('TryUpgrade("'+$old+'")')
    if($result -notlike 'PASS: ?*'){throw ('Upgrade from schema '+$previous+' failed: '+$result)}
    $copy=$result.Substring(6)
    if((Get-Schema $old) -ne $current){throw 'The upgraded data file has the wrong schema version'}
    if((Get-FileHash -LiteralPath $copy).Hash -ne $oldHash){throw 'The copy from before the upgrade is not the data file as it was'}
    $upgraded=Get-SchemaText $old -NoRows;$fresh=Get-SchemaText $BuildPath -NoRows
    $difference=@(Compare-Object $fresh.Lines $upgraded.Lines)
    if($difference.Count){throw ('Upgraded schema '+$previous+' differs from a fresh build: '+(($difference|Select-Object -First 4|ForEach-Object {$_.SideIndicator+' '+$_.InputObject}) -join ' | '))}
    Write-Output ('PASS: a schema '+$previous+' data file upgraded to '+$current+' matches a fresh build ('+$fresh.Lines.Count+' schema lines); the copy from before is kept beside it')
    if([string]$app.Eval('TryUpgrade("'+$old+'")') -ne 'PASS: '){throw 'A current data file was upgraded again'}

    #A step that fails after changing the file: the file is put back from the copy.
    $broken=Join-Path $folder 'failing upgrade data.accdb';New-SchemaFixture $previous $broken
    #A rule that refuses the new version number makes the last statement of the step fail.
    $db=$engine.OpenDatabase($broken);try{$db.TableDefs.Item('AppMetadata').ValidationRule='[SchemaVersion]<='+$previous}finally{$db.Close()}
    $brokenHash=(Get-FileHash -LiteralPath $broken).Hash
    $result=[string]$app.Eval('TryUpgrade("'+$broken+'")')
    if($result -notlike 'REJECTED:*put back as it was*'){throw ('A failed upgrade was not put back: '+$result)}
    if((Get-FileHash -LiteralPath $broken).Hash -ne $brokenHash -or (Get-Schema $broken) -ne $previous){throw 'The data file differs after a failed upgrade'}
    Write-Output 'PASS: an upgrade that fails part way puts the data file back exactly as it was'

    #Refused without a change: newer, too old, and open in another session.
    $newer=Join-Path $folder 'newer data.accdb';Copy-Item -LiteralPath $old -Destination $newer;Invoke-Dao $newer @('UPDATE AppMetadata SET SchemaVersion='+($current+1))
    $result=[string]$app.Eval('TryUpgrade("'+$newer+'")');if($result -notlike 'REJECTED:*newer than this version*'){throw ('Newer data file: '+$result)}
    $floor=[int]$app.Eval('UpgradeFloorValue()')
    $older=Join-Path $folder 'too old data.accdb';Copy-Item -LiteralPath $copy -Destination $older;Invoke-Dao $older @('UPDATE AppMetadata SET SchemaVersion='+($floor-1))
    $olderHash=(Get-FileHash -LiteralPath $older).Hash
    $result=[string]$app.Eval('TryUpgrade("'+$older+'")');if($result -notlike 'REJECTED:*older than this version can upgrade*'){throw ('Too old data file: '+$result)}
    if((Get-FileHash -LiteralPath $older).Hash -ne $olderHash){throw 'A refused data file was changed'}
    $busy=Join-Path $folder 'busy data.accdb';Copy-Item -LiteralPath $copy -Destination $busy;$busyHash=(Get-FileHash -LiteralPath $busy).Hash
    $holder=$engine.OpenDatabase($busy)
    $result=[string]$app.Eval('TryUpgrade("'+$busy+'")')
    $holder.Close();$holder=$null
    if($result -notlike 'REJECTED:*open in another Swift NAVFIT session*'){throw ('Data file in use: '+$result)}
    if((Get-FileHash -LiteralPath $busy).Hash -ne $busyHash -or @(Get-ChildItem -LiteralPath $folder -Filter 'busy data before schema*').Count){throw 'A data file in use was changed or copied'}
    $problem=[string]$app.Eval('DataFileProblem("'+$restore+'")')
    if($problem -notlike 'That file is a complete Swift NAVFIT workspace*'){throw ('A complete workspace was offered as a data file: '+$problem)}
    Write-Output 'PASS: newer, too old, in-use data files and complete workspaces are refused and left unchanged'
}finally{if($null -ne $holder){$holder.Close()};Stop-Front $app;$app=$null}

#At startup: a front end whose data file is one schema older upgrades it before anything
#else, keeps the user's data and keeps a copy of the file from before. The older file is
#made from the user's data file by removing what the previous schema did not have.
$front=Join-Path $folder 'front end.accdb'
$fixture=Get-SchemaText $copy -NoRows
$keep=@($fixture.Lines|Where-Object {$_ -like 'TABLE *'}|ForEach-Object {($_ -split ' ')[1]})
$downgrade=@()
$db=$engine.OpenDatabase($moved)
try{
    foreach($t in @($db.TableDefs|Where-Object {$_.Name -notlike 'MSys*' -and $_.Name -notlike '~*'}|ForEach-Object {$_.Name})){
        if($keep -notcontains $t){
            foreach($r in @($db.Relations|Where-Object {$_.Table -eq $t -or $_.ForeignTable -eq $t}|ForEach-Object {$_.Name})){$db.Relations.Delete($r)}
            $downgrade+=('DROP TABLE ['+$t+']')
        }else{
            foreach($f in @($db.TableDefs.Item($t).Fields|ForEach-Object {$_.Name})){if(-not @($fixture.Lines|Where-Object {$_ -like ($t+' F '+$f+' *')}).Count){$downgrade+=('ALTER TABLE ['+$t+'] DROP COLUMN ['+$f+']')}}
        }
    }
}finally{$db.Close()}
Invoke-Dao $moved ($downgrade+@('UPDATE AppMetadata SET SchemaVersion='+$previous))
$movedHash=(Get-FileHash -LiteralPath $moved).Hash
try{
    $app=Start-Front
    if((Get-Schema $moved) -ne $current){throw 'Startup did not upgrade the older data file'}
    if($app.Eval('DCount("*","Report")') -ne 1 -or $app.Eval('DLookup("CommandName","Workspace")') -ne 'SYNTHETIC DATA FILE COMMAND'){throw 'The upgrade lost the saved data'}
    $copies=@(Get-ChildItem -LiteralPath $folder -Filter ('moved workspace data before schema '+$current+' *.accdb'))
    if($copies.Count -ne 1 -or (Get-FileHash -LiteralPath $copies[0].FullName).Hash -ne $movedHash){throw 'No exact copy of the data file from before the upgrade'}
    $problem=[string]$app.Eval('DataLinkProblem()');if($problem -ne ''){throw ('Data file not ready after the upgrade: '+$problem)}
}finally{Stop-Front $app;$app=$null}
Write-Output ('PASS: at startup a schema '+$previous+' data file ('+$downgrade.Count+' changes back) is upgraded, with the data kept and a copy beside it')

[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($engine)
if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $buildHash){throw 'The build was changed'}
Write-Output 'PASS: a moved data file is reported at startup and works again once chosen'
