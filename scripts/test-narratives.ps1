param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$sourceHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$scratch=Join-Path $root ('build\tests\narrative-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'NarrativeHarness',(Join-Path $root 'tests\NarrativeHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestNarratives()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    $app.CloseCurrentDatabase();$app.OpenCurrentDatabase($scratch,$true)
    $db=$app.CurrentDb()
    $r=$db.OpenRecordset("SELECT RevisionID FROM ReportRevision AS R WHERE (SELECT Count(*) FROM RevisionValue AS V WHERE V.RevisionID=R.RevisionID AND V.FieldKey='OfficialComments')<>1",4)
    if(-not $r.EOF){throw 'A saved revision is missing its unique narrative snapshot.'}
    $r.Close();$db.Close();$app.CloseCurrentDatabase()
    @{Status=$result;Reopened=$true;BuildSHA256=$sourceHash;TestDatabase=$scratch}|ConvertTo-Json|Set-Content (Join-Path $root 'build\tests\narrative-results.json')
    Write-Output $result
}finally{
    if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $sourceHash){throw 'Narrative test changed its source artifact'}
}
