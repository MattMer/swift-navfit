param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/ux-surface-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try{
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'UXSurfaceHarness',(Join-Path $root 'tests/UXSurfaceHarness.bas'))
    $result=$app.Eval('TestUXSurface()')
    if(-not $result.StartsWith('PASS:')){throw $result}
    Write-Output $result
}finally{if($null -ne $app){$app.Quit(2)}}
