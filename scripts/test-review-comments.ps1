param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/comments-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'ReviewCommentsHarness',(Join-Path $root 'tests/ReviewCommentsHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestReviewComments()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    $app.CloseCurrentDatabase(); $app.OpenCurrentDatabase($scratch,$true)
    if(-not $app.Eval('OpenReviewComments(DMin("ReportID","ReviewComment"))')) {throw 'Comments reopen failed'}
    $db=$app.CurrentDb(); $counts=@{}
    foreach($name in @('ReviewComment','ReviewCommentEvent')) {$r=$db.OpenRecordset('SELECT Count(*) FROM ['+$name+']'); $counts[$name]=$r.Fields.Item(0).Value; $r.Close(); if($counts[$name] -eq 0){throw 'History did not persist'}}
    $db.Close(); $app.CloseCurrentDatabase()
    @{Status=$result;Counts=$counts;Reopened=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/review-comments-results.json')
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
