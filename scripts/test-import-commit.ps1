param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$source=Join-Path $root 'TEST_DB.accdb'
$hash=(Get-FileHash -LiteralPath $source).Hash
$scratch=Join-Path $root ('build\tests\import-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'ImportHarness',(Join-Path $root 'tests\ImportHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $native=$app.Eval('TryHash("'+$source+'")')
    if($native -ne $hash){throw ('Native hash differs: '+$native)}
    $result=$app.Eval('TestImports("'+$source+'")')
    if(-not $result.StartsWith('PASS:')){throw $result}
    $app.CloseCurrentDatabase();$app.OpenCurrentDatabase($scratch,$true)
    $db=$app.CurrentDb();$r=$db.OpenRecordset('SELECT Count(*) FROM ImportRawValue')
    if($r.Fields.Item(0).Value -ne 2472){throw 'Provenance did not persist.'};$r.Close();$db.Close();$app.CloseCurrentDatabase()
    if((Get-FileHash -LiteralPath $source).Hash -ne $hash){throw 'Original source changed.'}
    @{Status=$result;Reopened=$true;SourceUnchanged=$true;NativeSHA256Verified=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch}|ConvertTo-Json|Set-Content (Join-Path $root 'build\tests\import-commit-results.json')
    Write-Output $result
}finally{if($null -ne $app){$app.Quit(2)}}
