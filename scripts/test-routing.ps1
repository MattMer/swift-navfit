param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/routing-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'RoutingHarness',(Join-Path $root 'tests/RoutingHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestRouting()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    $app.CloseCurrentDatabase(); $app.OpenCurrentDatabase($scratch,$true)
    #A reopened workspace asks for this session's contributor before any change (C03).
    if(-not $app.Eval('CurrentProject.AllForms("frmContributorStart").IsLoaded') -or $app.Eval('HasActiveContributor()')){throw 'Reopen did not ask for a contributor'}
    $app.Forms.Item('frmContributorStart').Controls.Item('cboContributor').Value=[int]$app.Eval('DLookup("ActiveProfileID","Workspace","WorkspaceID=1")')
    if(-not $app.Eval('AcceptContributorStartup()')){throw 'Contributor choice after reopen failed'}
    if(-not $app.Eval('OpenMyWork()')) {throw 'My Work reopen failed'}
    $db=$app.CurrentDb(); $counts=@{}
    foreach($name in @('ReviewTask','ReviewTaskEvent')) {$r=$db.OpenRecordset('SELECT Count(*) FROM ['+$name+']'); $counts[$name]=$r.Fields.Item(0).Value; $r.Close(); if($counts[$name] -eq 0){throw 'Routing did not persist'}}
    $db.Close(); $app.CloseCurrentDatabase()
    @{Status=$result;Counts=$counts;Reopened=$true;BuildSHA256=(Get-FileHash -LiteralPath $BuildPath).Hash;TestDatabase=$scratch} | ConvertTo-Json | Set-Content (Join-Path $root 'build/tests/routing-results.json')
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
