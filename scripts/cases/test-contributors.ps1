param([Parameter(Mandatory=$true)][string]$BuildPath)
#P06 contributors from the directory, deleting contributors from routes and
#reviews, and report people fields filled from directory members.
$ErrorActionPreference='Stop'
$root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$scratch=Join-Path $root ('build/tests/contributors-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try {
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1; $app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'ContributorHarness',(Join-Path $root 'tests/ContributorHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestContributors()')
    if(-not $result.StartsWith('PASS:')) {throw $result}
    Write-Output $result
} finally {if($null -ne $app){$app.Quit(2)}}
