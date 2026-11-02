param([Parameter(Mandatory=$true)][string]$BuildPath)
#F02 roster snapshots and departures: PRD window, absence from a complete roster,
#partial rosters, returns, matching, rank updates and confirm/dismiss.
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/roster-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    foreach($opened in @($app.Forms|ForEach-Object {$_.Name})){$app.DoCmd.Close(2,$opened,2)}
    $app.LoadFromText(5,'RosterHarness',(Join-Path $root 'tests/RosterHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestRoster()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
