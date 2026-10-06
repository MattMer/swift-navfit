param([Parameter(Mandatory=$true)][string]$BuildPath)
$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$scratch=Join-Path $root ('build/tests/ux-responsive-identity-'+[guid]::NewGuid()+'.accdb')
Copy-Item -LiteralPath $BuildPath -Destination $scratch
$app=$null
try{
    $app=New-Object -ComObject Access.Application
    $app.AutomationSecurity=1;$app.OpenCurrentDatabase($scratch,$true)
    $app.LoadFromText(5,'UXResponsiveIdentityHarness',(Join-Path $root 'tests/UXResponsiveIdentityHarness.bas'))
    $app.DoCmd.RunCommand(126)
    $result=$app.Eval('TestUXResponsiveIdentity()')
    Write-Output $result
    if(-not $result.StartsWith('PASS:')){throw $result}
}finally{if($null -ne $app){try{[void]$app.Eval('UXLeaveEditor("Discard")')}catch{};$app.Quit(2)}}
