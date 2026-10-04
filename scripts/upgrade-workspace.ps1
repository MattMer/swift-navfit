param(
    [Parameter(Mandatory=$true)][string]$SourcePath,
    [string]$TemplatePath=(Join-Path $PSScriptRoot '..\output\Swift_NAVFIT_0.14.0-reporting-senior.accdb'),
    [Parameter(Mandatory=$true)][string]$DestinationPath
)
$ErrorActionPreference='Stop'
$SourcePath=[IO.Path]::GetFullPath($SourcePath)
$TemplatePath=[IO.Path]::GetFullPath($TemplatePath)
$DestinationPath=[IO.Path]::GetFullPath($DestinationPath)
if(Test-Path -LiteralPath $DestinationPath){throw 'Destination exists; refusing to overwrite.'}
$before=(Get-FileHash -LiteralPath $SourcePath).Hash
$work=Join-Path (Split-Path $DestinationPath -Parent) ('upgrade-'+[guid]::NewGuid()+'.accdb')
[void](New-Item -ItemType Directory -Force -Path (Split-Path $DestinationPath -Parent))
Copy-Item -LiteralPath $TemplatePath -Destination $work
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1 # Known target VBA only; source opened read-only via DAO, without macros.
    $app.OpenCurrentDatabase($work,$true)
    $app.LoadFromText(5,'modMigration',(Join-Path $PSScriptRoot '..\tests\modMigration.bas'))
    $result=$app.Eval('MigrateWorkspace("'+$SourcePath.Replace('"','""')+'")')
    if(-not $result.StartsWith('PASS:')){throw $result}
    $app.DoCmd.DeleteObject(5,'modMigration')
    $app.DoCmd.RunCommand(126)
    $app.CloseCurrentDatabase()
    $app.OpenCurrentDatabase($work,$true)
    $db=$app.CurrentDb();$counts=[ordered]@{}
    foreach($name in @('Workspace','UserProfile','Report','ReportRevision','RevisionValue','RevisionChange','ReportEvent','Folder','FolderMove','ImportReceipt','ImportRawValue','RecoveryDraft','OperationJournal','NarrativeSnippet','LibraryPreference','FolderBatchChange','ReviewComment','ReviewCommentEvent','ReviewTask','ReviewTaskEvent')){
        $r=$db.OpenRecordset('SELECT Count(*) FROM ['+$name+']');$counts[$name]=$r.Fields.Item(0).Value;$r.Close()
    }
    $db.Close();$app.CloseCurrentDatabase()
    if((Get-FileHash -LiteralPath $SourcePath).Hash -ne $before){throw 'Source hash changed.'}
    Move-Item -LiteralPath $work -Destination $DestinationPath
    $evidence=@{Status=$result;SourceSHA256=$before;Destination=$DestinationPath;Counts=$counts;Reopened=$true}
    $evidence|ConvertTo-Json -Depth 5|Set-Content -LiteralPath ($DestinationPath+'.migration.json') -Encoding utf8
    $evidence|ConvertTo-Json -Depth 5
}finally{
    if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
