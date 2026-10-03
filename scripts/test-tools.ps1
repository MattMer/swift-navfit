param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$source=Join-Path $root 'TEST_DB.accdb'
$before=(Get-FileHash -LiteralPath $source).Hash
$scratch=Join-Path $root ('build\tests\tools-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'ToolsHarness',(Join-Path $root 'tests\ToolsHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestTools("'+$source+'")')
    if(-not $result.StartsWith('PASS:')){throw $result}
    $app.CloseCurrentDatabase();$app.OpenCurrentDatabase($scratch,$true)
    $db=$app.CurrentDb();$counts=@{}
    foreach($name in @('NarrativeSnippet','LibraryPreference','OperationJournal','FolderBatchChange')){$r=$db.OpenRecordset("SELECT Count(*) FROM [$name]");$counts[$name]=$r.Fields.Item(0).Value;$r.Close();if($counts[$name] -eq 0){throw "$name did not persist"}}
    $db.Close();$app.CloseCurrentDatabase()
    if((Get-FileHash -LiteralPath $source).Hash -ne $before){throw 'Source changed.'}
    @{Status=$result;Reopened=$true;SourceUnchanged=$true;Counts=$counts;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $root 'build\tests\tools-results.json')
    Write-Output $result
}finally{if($null -ne $app){$app.Quit(2)}}
