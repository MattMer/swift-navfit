param([Parameter(Mandatory=$true)][string]$BuildPath)
#F01 shared-draft review chain: holder saves, forward, send back, pull back,
#override flags, Members-grid logging and group-route membership.
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/route-chain-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'RouteChainHarness',(Join-Path $root 'tests/RouteChainHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestRouteChain()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
