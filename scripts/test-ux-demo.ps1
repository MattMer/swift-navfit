param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$sourceHash=(Get-FileHash -LiteralPath $BuildPath).Hash
$scratch=Join-Path $root ('build/tests/ux-demo-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try{
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'UXDemoSeed',(Join-Path $root 'tests/UXDemoSeed.bas'))
    $app.LoadFromText(5,'UXDemoHarness',(Join-Path $root 'tests/UXDemoHarness.bas'))
    $result=$app.Eval('TestUXDemo()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    Write-Output $result
}finally{if($null -ne $app){$app.Quit(2)}}
if((Get-FileHash -LiteralPath $BuildPath).Hash -ne $sourceHash){throw 'Demo source was modified'}
