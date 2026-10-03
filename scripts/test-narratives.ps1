param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
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
    $db=$app.CurrentDb();$r=$db.OpenRecordset('SELECT Count(*) FROM RevisionValue WHERE FieldKey=''OfficialComments''')
    if($r.Fields.Item(0).Value -ne 21){throw 'Narrative snapshots did not persist.'};$r.Close();$db.Close();$app.CloseCurrentDatabase()
    @{Status=$result;Reopened=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch}|ConvertTo-Json|Set-Content (Join-Path $root 'build\tests\narrative-results.json')
    Write-Output $result
}finally{if($null -ne $app){$app.Quit(2)}}
