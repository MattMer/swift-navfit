param([Parameter(Mandatory=$true)][string]$BuildPath)
# Narratives suite: narrative persistence and the one-snapshot-per-revision check after
# reopening; narrative layout; NavyWriter snippets with pending and saved values.
$ErrorActionPreference='Stop';$root=Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'suite-common.ps1')
$buildHash=(Get-FileHash -LiteralPath $BuildPath).Hash

Invoke-SuiteCase 'narratives' {
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
        Write-Output $result
    }finally{
        if($null -ne $app){$app.Quit(2);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    }
}
Invoke-HarnessCase 'narrative-layout' $BuildPath 'NarrativeLayoutHarness' 'TestNarrativeLayout()'
Invoke-HarnessCase 'navywriter' $BuildPath 'UXNavyWriterHarness' 'TestUXNavyWriter()' -After {param($app) [void]$app.Eval('UXLeaveEditor("Discard")')}

Complete-SuiteRun 'narratives' $BuildPath $buildHash
